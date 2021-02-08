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

	"github.com/pbnjay/grate/xls/cfb"
	"github.com/pbnjay/grate/xls/crypto"
)

type WorkBook struct {
	filename string
	ctx      context.Context
	doc      cfb.Document

	h        *header
	sheets   []*boundSheet
	codepage uint16
	dateMode uint16
	strings  []string

	password   string
	substreams [][]*rec

	fpos          int64
	pos2substream map[int64]int

	decryptors map[int]crypto.Decryptor
}

func Open(ctx context.Context, filename string) (*WorkBook, error) {
	doc, err := cfb.Open(filename)
	if err != nil {
		return nil, err
	}

	b := &WorkBook{
		filename: filename,
		ctx:      ctx,
		doc:      doc,

		pos2substream: make(map[int64]int, 10),
	}

	rdr, err := doc.Open("Workbook")
	if err != nil {
		return nil, err
	}
	//br := bufio.NewReader(rdr)
	err = b.loadFromStream(rdr)
	return b, err
}

func (b *WorkBook) loadFromStream(r io.Reader) error {
	b.decryptors = make(map[int]crypto.Decryptor)
	b.h = &header{}
	substr := -1
	nr, err := b.nextRecord(r)
	for err == nil {
		if nr.RecType == RecTypeBOF {
			substr++
			b.substreams = append(b.substreams, []*rec{})
			b.pos2substream[b.fpos] = substr
		}
		b.fpos += int64(4 + len(nr.Data))

		if nr.RecType == RecTypeFilePass {
			etype := binary.LittleEndian.Uint16(nr.Data)
			switch etype {
			case 1:
				b.decryptors[substr], err = crypto.NewBasicRC4(nr.Data[2:])
				if err != nil {
					log.Println("xls: rc4 encryption failed to set up", err)
					return err
				}
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
		log.Printf("Processing substream %d/%d (%d records)", ss, len(b.substreams), len(records))

		if dec, ok := b.decryptors[ss]; ok {
			log.Printf("Decrypting substream...")

			dec.Reset()
			var head [4]byte
			for _, nr := range records {
				binary.LittleEndian.PutUint16(head[:], uint16(nr.RecType))
				binary.LittleEndian.PutUint16(head[2:], nr.RecSize)

				// send the record for decryption
				dec.Write(head[:])
				dec.Write(nr.Data)
			}
			dec.Flush()

			newrecset := make([]*rec, 0, len(records))
			for _, nr := range records {
				dec.Read(head[:]) // discard 4 byte header

				dr := &rec{
					RecType: nr.RecType,
					RecSize: nr.RecSize,
					Data:    make([]byte, int(nr.RecSize)),
				}
				dec.Read(dr.Data)

				switch nr.RecType {
				case RecTypeBOF, RecTypeFilePass, RecTypeUsrExcl, RecTypeFileLock, RecTypeInterfaceHdr, RecTypeRRDInfo, RecTypeRRDHead:
					// keep original data
					copy(dr.Data, nr.Data)
				case RecTypeBoundSheet8:
					// copy the position un-decrypted
					copy(dr.Data[:4], nr.Data)
				default:
					// apply decryption
				}

				newrecset = append(newrecset, dr)
			}

			b.substreams[ss] = newrecset
			records = newrecset
		}

		for i, nr := range records {
			var bb io.Reader = bytes.NewReader(nr.Data)

			switch nr.RecType {
			case RecTypeSST:
				//log.Println(i, nr.RecType)

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
				log.Println("End Of Stream")

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
					break
				}

			case RecTypeCodePage:
				//log.Println(i, nr.RecType)
				err = binary.Read(bb, binary.LittleEndian, &b.codepage)
				if err != nil {
					return err
				}

			case RecTypeDate1904:
				//log.Println(i, nr.RecType)
				err = binary.Read(bb, binary.LittleEndian, &b.dateMode)
				if err != nil {
					return err
				}
			case RecTypeBoundSheet8:
				//log.Println(i, nr.RecType)
				bs := &boundSheet{}
				err = binary.Read(bb, binary.LittleEndian, &bs.Position)
				if err != nil {
					log.Println("fail1", err)
					return err
				}

				err = binary.Read(bb, binary.LittleEndian, &bs.HiddenState)
				if err != nil {
					log.Println("fail1", err)
					return err
				}
				err = binary.Read(bb, binary.LittleEndian, &bs.SheetType)
				if err != nil {
					log.Println("fail1", err)
					return err
				}

				bs.Name, err = decodeShortXLUnicodeString(bb)
				if err != nil {
					log.Println("fail2", err)
					return err
				}
				b.sheets = append(b.sheets, bs)
				log.Println("SHEET", bs.Name, "at pos", bs.Position)
			default:
				//log.Println(i, "SKIPPED", nr.RecType)
			}
		}
	}

	return err
}

var errSkipped = errors.New("xls: skipped record type")

func (b *WorkBook) nextRecord(r io.Reader) (*rec, error) {
	var rt recordType
	var rs uint16
	err := binary.Read(r, binary.LittleEndian, &rt)
	if err != nil {
		return nil, err
	}

	err = binary.Read(r, binary.LittleEndian, &rs)
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
