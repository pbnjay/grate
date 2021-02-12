// Package xls implements the Microsoft Excel Binary File Format (.xls) Structure.
// More specifically, it contains just enough detail to extract cell contents,
// data types, and last-calculated formula values. In particular, it does NOT
// implement formatting or formula calculations.
package xls

// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06

import (
	"bytes"
	"context"
	"encoding/binary"
	"errors"
	"io"
	"log"

	"github.com/pbnjay/grate"
	"github.com/pbnjay/grate/commonxl"
	"github.com/pbnjay/grate/xls/cfb"
	"github.com/pbnjay/grate/xls/crypto"
)

var _ = grate.Register("xls", 1, Open)

type WorkBook struct {
	filename string
	ctx      context.Context
	doc      *cfb.Document

	prot     bool
	h        *header
	sheets   []*boundSheet
	codepage uint16
	dateMode uint16
	strings  []string

	password   string
	substreams [][]*rec

	fpos          int64
	pos2substream map[int64]int

	nfmt commonxl.Formatter
	xfs  []uint16
}

func (b *WorkBook) IsProtected() bool {
	return b.prot
}

func Open(filename string) (grate.Source, error) {
	doc, err := cfb.Open(filename)
	if err != nil {
		return nil, err
	}

	b := &WorkBook{
		filename: filename,
		doc:      doc,

		pos2substream: make(map[int64]int, 16),
		xfs:           make([]uint16, 0, 128),
	}

	rdr, err := doc.Open("Workbook")
	if err != nil {
		return nil, grate.WrapErr(err, grate.ErrNotInFormat)
	}
	err = b.loadFromStream(rdr)
	return b, err
}

func (b *WorkBook) loadFromStream(r io.ReadSeeker) error {
	return b.loadFromStream2(r, false)
}

func (b *WorkBook) loadFromStreamWithDecryptor(r io.ReadSeeker, dec crypto.Decryptor) error {
	if grate.Debug {
		log.Println("  Decrypting xls stream with standard RC4")
	}
	_, err := r.Seek(0, io.SeekStart)
	if err != nil {
		log.Println("xls: dec-seek1 failed")
		return err
	}

	zeros := [8224]byte{}

	type overlay struct {
		Pos int64

		RecType   recordType
		DataBytes uint16
		Data      []byte // NB len() not necessarily = DataBytes
	}
	replaceBlocks := []overlay{}

	obuf := &bytes.Buffer{}
	for err == nil {
		o := overlay{}
		o.Pos, _ = r.Seek(0, io.SeekCurrent)

		err = binary.Read(r, binary.LittleEndian, &o.RecType)
		if err != nil {
			if err == io.EOF {
				continue
			}
			log.Println("xls: dec-read1 failed")
			return err
		}

		err = binary.Read(r, binary.LittleEndian, &o.DataBytes)
		if err != nil {
			log.Println("xls: dec-read2 failed")
			return err
		}

		// copy to output and decryption stream
		binary.Write(dec, binary.LittleEndian, o.RecType)
		binary.Write(dec, binary.LittleEndian, o.DataBytes)
		tocopy := int(o.DataBytes)

		switch o.RecType {
		case RecTypeBOF, RecTypeFilePass, RecTypeUsrExcl, RecTypeFileLock, RecTypeInterfaceHdr, RecTypeRRDInfo, RecTypeRRDHead:
			// copy original data into output
			o.Data = make([]byte, o.DataBytes)
			_, err = io.ReadFull(r, o.Data)
			if err != nil {
				log.Println("FAIL err", err)
			}
			dec.Write(zeros[:int(o.DataBytes)])
			tocopy = 0

		case RecTypeBoundSheet8:
			// copy 32-bit position to output
			o.Data = make([]byte, 4)
			_, err = io.ReadFull(r, o.Data)
			if err != nil {
				log.Println("FAIL err", err)
			}
			dec.Write(zeros[:4])
			tocopy -= 4
		}

		if tocopy > 0 {
			_, err = io.CopyN(dec, r, int64(tocopy))
		}
		replaceBlocks = append(replaceBlocks, o)
	}
	dec.Flush()
	io.Copy(obuf, dec)

	alldata := obuf.Bytes()
	for _, o := range replaceBlocks {
		offs := int(o.Pos)
		binary.LittleEndian.PutUint16(alldata[offs:], uint16(o.RecType))
		binary.LittleEndian.PutUint16(alldata[offs+2:], uint16(o.DataBytes))
		if len(o.Data) > 0 {
			offs += 4
			copy(alldata[offs:], o.Data)
		}
	}

	return b.loadFromStream2(bytes.NewReader(alldata), true)
}

func (b *WorkBook) loadFromStream2(r io.ReadSeeker, isDecrypted bool) error {
	b.h = &header{}
	substr := -1
	nestedBOF := 0
	b.substreams = b.substreams[:0]
	b.pos2substream = make(map[int64]int, 10)
	b.fpos = 0
	nr, err := b.nextRecord(r)
	for err == nil {
		switch nr.RecType {
		case RecTypeEOF:
			nestedBOF--
		case RecTypeBOF:
			// when substreams are nested, keep them in the same grouping
			if nestedBOF == 0 {
				substr = len(b.substreams)
				b.substreams = append(b.substreams, []*rec{})
				b.pos2substream[b.fpos] = substr
			}
			nestedBOF++
		}
		b.fpos += int64(4 + len(nr.Data))

		if nr.RecType == RecTypeFilePass && !isDecrypted {
			etype := binary.LittleEndian.Uint16(nr.Data)
			switch etype {
			case 1:
				dec, err := crypto.NewBasicRC4(nr.Data[2:])
				if err != nil {
					log.Println("xls: rc4 encryption failed to set up", err)
					return err
				}
				return b.loadFromStreamWithDecryptor(r, dec)
			case 2, 3, 4:
				log.Println("need Crypto API RC4 decryptor")
				return errors.New("xls: unsupported Crypto API encryption method")
			default:
				return errors.New("xls: unsupported encryption method")
			}
		}

		b.substreams[substr] = append(b.substreams[substr], nr)
		nr, err = b.nextRecord(r)
	}
	if err == io.EOF {
		err = nil
	}
	if err != nil {
		return err
	}

	for ss, records := range b.substreams {
		if grate.Debug {
			log.Printf("  Processing substream %d/%d (%d records)", ss, len(b.substreams), len(records))
		}
		for i, nr := range records {
			var bb io.Reader = bytes.NewReader(nr.Data)

			switch nr.RecType {
			case RecTypeSST:
				recSet := []*rec{nr}

				lastIndex := i
				for len(records) > (lastIndex+1) && records[lastIndex+1].RecType == RecTypeContinue {
					lastIndex++
					recSet = append(recSet, records[lastIndex])
				}

				b.strings, err = parseSST(recSet)
				if err != nil {
					return err
				}

			case RecTypeContinue:
				// no-op (used above)
			case RecTypeEOF:
				// done

			case RecTypeBOF:
				err = binary.Read(bb, binary.LittleEndian, b.h)
				if err != nil {
					return err
				}

				if b.h.Version != 0x0600 {
					return errors.New("xls: invalid file version")
				}
				if b.h.RupYear != 0x07CC && b.h.RupYear != 0x07CD {
					return errors.New("xls: unsupported biff version")
				}
				if b.h.DocType != 0x0005 && b.h.DocType != 0x0010 {
					// we only support the workbook or worksheet substreams
					log.Println("xls: unsupported document type")
					//break
				}

			case RecTypeCodePage:
				err = binary.Read(bb, binary.LittleEndian, &b.codepage)
				if err != nil {
					return err
				}

			case RecTypeDate1904:
				err = binary.Read(bb, binary.LittleEndian, &b.dateMode)
				if err != nil {
					return err
				}

			case RecTypeFormat:
				var fmtNo uint16
				err = binary.Read(bb, binary.LittleEndian, &fmtNo)
				formatStr, err := decodeXLUnicodeString(bb)
				if err != nil {
					log.Println("fail2", err)
					return err
				}
				b.nfmt.Add(fmtNo, formatStr)

			case RecTypeXF:
				var x, fmtNo uint16
				err = binary.Read(bb, binary.LittleEndian, &x) // ignore font
				err = binary.Read(bb, binary.LittleEndian, &fmtNo)
				b.xfs = append(b.xfs, fmtNo)

			case RecTypeBoundSheet8:
				bs := &boundSheet{}
				err = binary.Read(bb, binary.LittleEndian, &bs.Position)
				if err != nil {
					return err
				}

				err = binary.Read(bb, binary.LittleEndian, &bs.HiddenState)
				if err != nil {
					return err
				}
				err = binary.Read(bb, binary.LittleEndian, &bs.SheetType)
				if err != nil {
					return err
				}

				bs.Name, err = decodeShortXLUnicodeString(bb)
				if err != nil {
					return err
				}
				b.sheets = append(b.sheets, bs)
			default:
				if grate.Debug && ss == 0 {
					log.Println("    Unhandled record type:", nr.RecType, i)
				}
			}
		}
	}

	return err
}

func (b *WorkBook) nextRecord(r io.Reader) (*rec, error) {
	var rt recordType
	var rs uint16
	err := binary.Read(r, binary.LittleEndian, &rt)
	if err != nil {
		return nil, err
	}
	if rt == 0 {
		return nil, io.EOF
	}

	err = binary.Read(r, binary.LittleEndian, &rs)
	if rs > 8224 {
		return nil, errors.New("xls: invalid data format")
	}
	if err != nil {
		return nil, err
	}

	data := make([]byte, rs)
	_, err = io.ReadFull(r, data)
	if err != nil {
		return nil, err
	}

	ret := &rec{rt, rs, data}
	return ret, err
}
