// Package cfb implements the Microsoft Compound File Binary File Format.
package cfb

// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b
// Note for myself:
//   Storage = Directory
//   Stream = File

import (
	"bytes"
	"encoding/binary"
	"errors"
	"io"
	"io/ioutil"
	"log"
	"unicode/utf16"

	"github.com/pbnjay/grate"
)

const fullAssertions = true

const (
	secFree       uint32 = 0xFFFFFFFF // FREESECT
	secEndOfChain uint32 = 0xFFFFFFFE // ENDOFCHAIN
	secFAT        uint32 = 0xFFFFFFFD // FATSECT
	secDIFAT      uint32 = 0xFFFFFFFC // DIFSECT
	secReserved   uint32 = 0xFFFFFFFB
	secMaxRegular uint32 = 0xFFFFFFFA // MAXREGSECT
)

// Header of the Compound File MUST be at the beginning of the file (offset 0).
type header struct {
	Signature                    uint64      // Identification signature for the compound file structure, and MUST be set to the value 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1.
	ClassID                      [2]uint64   // Reserved and unused class ID that MUST be set to all zeroes (CLSID_NULL).
	MinorVersion                 uint16      // Version number for nonbreaking changes. This field SHOULD be set to 0x003E if the major version field is either 0x0003 or 0x0004.
	MajorVersion                 uint16      // Version number for breaking changes. This field MUST be set to either 0x0003 (version 3) or 0x0004 (version 4).
	ByteOrder                    uint16      // This field MUST be set to 0xFFFE. This field is a byte order mark for all integer fields, specifying little-endian byte order.
	SectorShift                  uint16      // This field MUST be set to 0x0009, or 0x000c, depending on the Major Version field. This field specifies the sector size of the compound file as a power of 2.
	MiniSectorShift              uint16      // This field MUST be set to 0x0006. This field specifies the sector size of the Mini Stream as a power of 2. The sector size of the Mini Stream MUST be 64 bytes.
	Reserved1                    [6]byte     // This field MUST be set to all zeroes.
	NumDirectorySectors          int32       // This integer field contains the count of the number of directory sectors in the compound file.
	NumFATSectors                int32       // This integer field contains the count of the number of FAT sectors in the compound file.
	FirstDirectorySectorLocation uint32      // This integer field contains the starting sector number for the directory stream.
	TransactionSignature         int32       // This integer field MAY contain a sequence number that is incremented every time the compound file is saved by an implementation that supports file transactions. This is the field that MUST be set to all zeroes if file transactions are not implemented.<1>
	MiniStreamCutoffSize         int32       // This integer field MUST be set to 0x00001000. This field specifies the maximum size of a user-defined data stream that is allocated from the mini FAT and mini stream, and that cutoff is 4,096 bytes. Any user-defined data stream that is greater than or equal to this cutoff size must be allocated as normal sectors from the FAT.
	FirstMiniFATSectorLocation   uint32      // This integer field contains the starting sector number for the mini FAT.
	NumMiniFATSectors            int32       // This integer field contains the count of the number of mini FAT sectors in the compound file.
	FirstDIFATSectorLocation     uint32      // This integer field contains the starting sector number for the DIFAT.
	NumDIFATSectors              int32       // This integer field contains the count of the number of DIFAT sectors in the compound file.
	DIFAT                        [109]uint32 // This array of 32-bit integer fields contains the first 109 FAT sector locations of the compound file.
}

type objectType byte

const (
	typeUnknown     objectType = 0x00
	typeStorage     objectType = 0x01
	typeStream      objectType = 0x02
	typeRootStorage objectType = 0x05
)

type directory struct {
	Name                   [32]uint16 // 32 utf16 characters
	NameByteLen            int16      // length of Name in bytes
	ObjectType             objectType
	ColorFlag              byte   // 0=red, 1=black
	LeftSiblingID          uint32 // stream ids
	RightSiblingID         uint32
	ChildID                uint32
	ClassID                [2]uint64 // GUID
	StateBits              uint32
	CreationTime           int64
	ModifiedTime           int64
	StartingSectorLocation int32
	StreamSize             uint64
}

func (d *directory) String() string {
	if (d.NameByteLen&1) == 1 || d.NameByteLen > 64 {
		return "<invalid utf16 string>"
	}
	r16 := utf16.Decode(d.Name[:int(d.NameByteLen)/2])
	// trim off null terminator
	return string(r16[:len(r16)-1])
}

// Document represents a Compound File Binary Format document.
type Document struct {
	// the entire file, loaded into memory
	data []byte

	// pre-parsed info
	header *header
	dir    []*directory

	// lookup tables for all the sectors
	fat     []uint32
	minifat []uint32

	ministreamstart uint32
	ministreamsize  uint32
}

func (d *Document) load(rx io.ReadSeeker) error {
	var err error
	d.data, err = ioutil.ReadAll(rx)
	if err != nil {
		return err
	}
	br := bytes.NewReader(d.data)

	h := &header{}
	err = binary.Read(br, binary.LittleEndian, h)
	if h.Signature != 0xe11ab1a1e011cfd0 {
		return grate.ErrNotInFormat // errors.New("ole2: invalid format")
	}
	if h.ByteOrder != 0xFFFE {
		return grate.ErrNotInFormat //errors.New("ole2: invalid format")
	}
	if fullAssertions {
		if h.ClassID[0] != 0 || h.ClassID[1] != 0 {
			return grate.ErrNotInFormat //errors.New("ole2: invalid CLSID")
		}
		if h.MajorVersion != 3 && h.MajorVersion != 4 {
			return errors.New("ole2: unknown major version")
		}
		if h.MinorVersion != 0x3E {
			log.Printf("WARNING MinorVersion = 0x%02x NOT 0x3E", h.MinorVersion)
			//return errors.New("ole2: unknown minor version")
		}

		for _, v := range h.Reserved1 {
			if v != 0 {
				return errors.New("ole2: reserved section is non-zero")
			}
		}
		if h.MajorVersion == 3 {
			if h.SectorShift != 9 {
				return errors.New("ole2: invalid sector size")
			}
			if h.NumDirectorySectors != 0 {
				return errors.New("ole2: version 3 does not support directory sectors")
			}
		}
		if h.MajorVersion == 4 {
			if h.SectorShift != 12 {
				return errors.New("ole2: invalid sector size")
			}
		}
		if h.MiniSectorShift != 6 {
			return errors.New("ole2: invalid mini sector size")
		}
		if h.MiniStreamCutoffSize != 0x00001000 {
			return errors.New("ole2: invalid mini sector cutoff")
		}
	}
	d.header = h

	numFATentries := (1 << (h.SectorShift - 2))
	le := binary.LittleEndian
	d.fat = make([]uint32, 0, numFATentries*int(1+d.header.NumFATSectors))
	d.minifat = make([]uint32, 0, numFATentries*int(1+h.NumMiniFATSectors))

	// step 1: read the DIFAT sector list
	for i := 0; i < 109; i++ {
		sid := h.DIFAT[i]
		if sid == secFree {
			break
		}
		offs := int64(1+sid) << int32(h.SectorShift)
		if offs >= int64(len(d.data)) {
			return errors.New("xls/cfb: unable to load file")
		}
		sector := d.data[offs:]
		for j := 0; j < numFATentries; j++ {
			sid2 := le.Uint32(sector)
			d.fat = append(d.fat, sid2)
			sector = sector[4:]
		}
	}
	if h.NumDIFATSectors > 0 {
		sid1 := h.FirstDIFATSectorLocation

		for sid1 != secEndOfChain {
			offs := int64(1+sid1) << int32(h.SectorShift)
			difatSector := d.data[offs:]

			for i := 0; i < numFATentries-1; i++ {
				sid2 := le.Uint32(difatSector)
				if sid2 == secFree || sid2 == secEndOfChain {
					difatSector = difatSector[4:]
					continue
				}

				offs := int64(1+sid2) << int32(h.SectorShift)
				if offs >= int64(len(d.data)) {
					return errors.New("xls/cfb: unable to load file")
				}
				sector := d.data[offs:]
				for j := 0; j < numFATentries; j++ {
					sid3 := le.Uint32(sector)
					d.fat = append(d.fat, sid3)
					sector = sector[4:]
				}

				difatSector = difatSector[4:]
			}
			// chain the next DIFAT sector
			sid1 = le.Uint32(difatSector)
		}
	}

	// step 2: read the mini FAT
	sid := h.FirstMiniFATSectorLocation
	for sid != secEndOfChain {
		offs := int64(1+sid) << int32(h.SectorShift)
		if offs >= int64(len(d.data)) {
			return errors.New("xls/cfb: unable to load file")
		}
		sector := d.data[offs:]
		for j := 0; j < numFATentries; j++ {
			sid = le.Uint32(sector)
			d.minifat = append(d.minifat, sid)
			sector = sector[4:]
		}

		if len(d.minifat) >= int(h.NumMiniFATSectors) {
			break
		}

		// chain the next mini FAT sector
		sid = le.Uint32(sector)
	}

	// step 3: read the Directory Entries
	err = d.buildDirs(br)

	return err
}

func (d *Document) buildDirs(br *bytes.Reader) error {
	h := d.header
	le := binary.LittleEndian

	// step 2: read the Directory
	sid := h.FirstDirectorySectorLocation
	offs := int64(1+sid) << int64(h.SectorShift)
	br.Seek(offs, io.SeekStart)

	for j := 0; j < 4; j++ {
		dirent := &directory{}
		binary.Read(br, le, dirent)
		if d.header.MajorVersion == 3 {
			// mask out upper 32bits
			dirent.StreamSize = dirent.StreamSize & 0xFFFFFFFF
		}

		switch dirent.ObjectType {
		case typeRootStorage:
			d.ministreamstart = uint32(dirent.StartingSectorLocation)
			d.ministreamsize = uint32(dirent.StreamSize)
		case typeStorage:
			//log.Println("got a storage? what to do now?")
		case typeStream:
			/*
				var freader io.Reader
				if dirent.StreamSize < uint64(d.header.MiniStreamCutoffSize) {
					freader = d.getMiniStreamReader(uint32(dirent.StartingSectorLocation), dirent.StreamSize)
				} else if dirent.StreamSize != 0 {
					freader = d.getStreamReader(uint32(dirent.StartingSectorLocation), dirent.StreamSize)
				}
			*/
		case typeUnknown:
			return nil
		}
		d.dir = append(d.dir, dirent)
	}

	return nil
}

func (d *Document) getStreamReader(sid uint32, size uint64) (io.ReadSeeker, error) {
	// NB streamData is a slice of slices of the raw data, so this is the
	// only allocation - for the (much smaller) list of sector slices
	streamData := make([][]byte, 1+(size>>d.header.SectorShift))

	x := 0
	secSize := int64(1) << int32(d.header.SectorShift)
	for sid != secEndOfChain && sid != secFree {
		offs := int64(1+sid) << int64(d.header.SectorShift)
		if offs > int64(len(d.data)) {
			return nil, errors.New("ole2: corrupt data format")
		}
		slice := d.data[offs : offs+secSize]
		if size < uint64(len(slice)) {
			slice = slice[:size]
			size = 0
		} else {
			size -= uint64(len(slice))
		}
		streamData[x] = slice
		if size == 0 {
			break
		}
		sid = d.fat[sid]
		x++
	}
	if size != 0 {
		return nil, errors.New("ole2: incomplete read")
	}

	return &SliceReader{Data: streamData}, nil
}

func (d *Document) getMiniStreamReader(sid uint32, size uint64) (io.ReadSeeker, error) {
	// TODO: move into a separate cache so we don't recalculate it each time
	fatStreamData := make([][]byte, 1+(d.ministreamsize>>d.header.SectorShift))

	// NB streamData is a slice of slices of the raw data, so this is the
	// only allocation - for the (much smaller) list of sector slices
	streamData := make([][]byte, 1+(size>>d.header.MiniSectorShift))

	x := 0
	fsid := d.ministreamstart
	fsize := uint64(d.ministreamsize)
	secSize := int64(1) << int64(d.header.SectorShift)
	for fsid != secEndOfChain && fsid != secFree {
		offs := int64(1+fsid) << int64(d.header.SectorShift)
		slice := d.data[offs : offs+secSize]
		if fsize < uint64(len(slice)) {
			slice = slice[:fsize]
			fsize = 0
		} else {
			fsize -= uint64(len(slice))
		}
		fatStreamData[x] = slice
		x++
		fsid = d.fat[fsid]
	}

	x = 0
	miniSecSize := int64(1) << int64(d.header.MiniSectorShift)
	for sid != secEndOfChain && sid != secFree {
		offs := int64(sid) << int64(d.header.MiniSectorShift)

		so, si := offs/secSize, offs%secSize
		data := fatStreamData[so]

		slice := data[si : si+miniSecSize]
		if size < uint64(len(slice)) {
			slice = slice[:size]
			size = 0
		} else {
			size -= uint64(len(slice))
		}
		streamData[x] = slice
		x++
		sid = d.minifat[sid]
	}

	return &SliceReader{Data: streamData}, nil
}
