package xls

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"log"
	"math"
	"time"
	"unicode/utf16"
)

func (b *WorkBook) Sheets() []string {
	res := make([]string, len(b.sheets))
	for i, s := range b.sheets {
		res[i] = s.Name
	}
	return res
}

func (b *WorkBook) Get(sheetName string) (*WorkSheet, error) {
	for _, s := range b.sheets {
		if s.Name == sheetName {
			ss := b.pos2substream[int64(s.Position)]
			ws := &WorkSheet{
				b: b, s: s, ss: ss,
			}
			return ws, ws.parse()
		}
	}
	return nil, errors.New("xls: sheet not found")
}

type WorkSheet struct {
	b   *WorkBook
	s   *boundSheet
	ss  int
	err error

	rows   []*row
	maxcol int

	iterRow int
}

type staticBlankType byte

const staticBlank staticBlankType = 0

func (staticBlankType) String() string {
	return ""
}

type row struct {
	// each value must be one of: int, float64, string, or time.Time
	cols []interface{}
}

func (s *WorkSheet) placeValue(rowIndex, colIndex int, val interface{}) {
	// ensure we always have a complete matrix
	if s.maxcol <= colIndex {
		s.maxcol = colIndex + 1
		for _, r := range s.rows {
			for len(r.cols) <= colIndex {
				r.cols = append(r.cols, staticBlank)
			}
		}
	}

	for len(s.rows) <= rowIndex {
		emptyRow := make([]interface{}, s.maxcol)
		for i := 0; i < s.maxcol; i++ {
			emptyRow[i] = staticBlank
		}
		s.rows = append(s.rows, &row{emptyRow})
	}

	s.rows[rowIndex].cols[colIndex] = val
}

func (s *WorkSheet) parse() error {
	for _, r := range s.b.substreams[s.ss] {
		bb := bytes.NewReader(r.Data)

		switch r.RecType {
		case RecTypeWindow2:
			opts := binary.LittleEndian.Uint16(r.Data)
			// right-to-left = 0x40, selected = 0x400
			log.Printf("sheet options: %x", opts)
		case RecTypeRow:
			row := &shRow{}
			binary.Read(bb, binary.LittleEndian, row)
			log.Printf("row spec: %+v", *row)
		case RecTypeBlank:
			var rowIndex, colIndex uint16
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &colIndex)
			log.Printf("blank spec: %d %d", rowIndex, colIndex)
		case RecTypeMulBlank:
			var rowIndex, firstCol uint16
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &firstCol)
			nrk := int((r.RecSize - 6) / 6)
			log.Printf("row blanks spec: %d %d %d", rowIndex, firstCol, nrk)
		case RecTypeMulRk:
			mr := &shMulRK{}
			nrk := int((r.RecSize - 6) / 6)
			binary.Read(bb, binary.LittleEndian, &mr.RowIndex)
			binary.Read(bb, binary.LittleEndian, &mr.FirstCol)
			mr.Values = make([]RkRec, nrk)
			for i := 0; i < nrk; i++ {
				rr := RkRec{}
				binary.Read(bb, binary.LittleEndian, &rr)
				mr.Values[i] = rr

				var rval interface{}
				if rr.Value.IsInteger() {
					rval = rr.Value.Int()
				} else {
					rval = rr.Value.Float64()
				}
				s.placeValue(int(mr.RowIndex), int(mr.FirstCol)+i, rval)
			}
			binary.Read(bb, binary.LittleEndian, &mr.LastCol)

			log.Printf("mulrow spec: %+v", *mr)

		case RecTypeNumber:
			var rowIndex, colIndex, ixfe uint16
			var xnum uint64
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &colIndex)
			binary.Read(bb, binary.LittleEndian, &ixfe)
			binary.Read(bb, binary.LittleEndian, &xnum)
			value := math.Float64frombits(xnum)
			s.placeValue(int(rowIndex), int(colIndex), value)
			log.Printf("Number spec: %d %d = %f", rowIndex, colIndex, value)

		case RecTypeRK:
			var rowIndex, colIndex uint16
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &colIndex)
			rr := RkRec{}
			binary.Read(bb, binary.LittleEndian, &rr)

			var rval interface{}
			if rr.Value.IsInteger() {
				rval = rr.Value.Int()
			} else {
				rval = rr.Value.Float64()
			}
			s.placeValue(int(rowIndex), int(colIndex), rval)
			log.Printf("RK spec: %d %d = %s", rowIndex, colIndex, rr.Value.String())

		case RecTypeFormula:
			var rowIndex, colIndex uint16
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &colIndex)

			log.Printf("formula spec: %d %d ~~ %+v", rowIndex, colIndex, r.Data)

		case RecTypeString:
			var charCount, flags uint16
			binary.Read(bb, binary.LittleEndian, &charCount)
			binary.Read(bb, binary.LittleEndian, &flags)
			s := ""
			if (flags & 1) == 0 {
				s = string(r.Data[4:])
			} else {
				us := make([]uint16, charCount)
				binary.Read(bb, binary.LittleEndian, us)
				s = string(utf16.Decode(us))
			}
			log.Printf("string spec:  = %s", s)

		case RecTypeLabelSst:
			var rowIndex, colIndex, ixfe uint16
			var sstIndex uint32
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &colIndex)
			binary.Read(bb, binary.LittleEndian, &ixfe)
			binary.Read(bb, binary.LittleEndian, &sstIndex)
			s.placeValue(int(rowIndex), int(colIndex), s.b.strings[sstIndex])
			log.Printf("SST spec: %d %d = [%d] %s", rowIndex, colIndex, sstIndex, s.b.strings[sstIndex])

		case RecTypeHLink:
			loc := &shRef8{}
			binary.Read(bb, binary.LittleEndian, loc)
			var x uint64
			binary.Read(bb, binary.LittleEndian, &x) // skip and discard classid
			binary.Read(bb, binary.LittleEndian, &x)
			var flags, slen uint32
			binary.Read(bb, binary.LittleEndian, &slen)
			if slen != 2 {
				log.Println("unknown hyperlink version")
				continue
			}
			str := "<hyperlink>"
			binary.Read(bb, binary.LittleEndian, &flags)
			if (flags & 0x10) != 0 {
				binary.Read(bb, binary.LittleEndian, &slen)
				us := make([]uint16, slen)
				binary.Read(bb, binary.LittleEndian, us)
				str = string(utf16.Decode(us))
			}

			// TODO: apply merge cell rules
			s.placeValue(int(loc.FirstRow), int(loc.FirstCol), str)
			log.Printf("hyperlink spec: %+v = %s", loc, str)

		case RecTypeMergeCells:
			var cmcs uint16
			binary.Read(bb, binary.LittleEndian, &cmcs)
			mcRefs := make([]shRef8, cmcs)
			binary.Read(bb, binary.LittleEndian, &mcRefs)
			log.Printf("MergeCells spec: %d records", cmcs)
			for j, mc := range mcRefs {
				log.Printf("    %d: %+v", j, mc)
			}

		default:
			log.Println("worksheet", r.RecType, r.RecSize)

		}
	}
	return nil
}

// Err returns the last error that occured.
func (s *WorkSheet) Err() error {
	return s.err
}

// Next advances to the next row of content.
// It MUST be called prior to any Scan().
func (s *WorkSheet) Next() bool {
	s.iterRow++
	return s.iterRow < len(s.rows)
}

func (s *WorkSheet) Strings() []string {
	currow := s.rows[s.iterRow]
	res := make([]string, len(currow.cols))
	for i, col := range currow.cols {
		res[i] = fmt.Sprint(col)
	}
	return res
}

// Scan extracts values from the row into the provided arguments
// Arguments must be pointers to one of 4 supported types:
//     int, float64, string, or time.Time
func (s *WorkSheet) Scan(args ...interface{}) error {
	currow := s.rows[s.iterRow]

	for i, a := range args {
		switch v := a.(type) {
		case *int:
			*v = currow.cols[i].(int)
		case *float64:
			*v = currow.cols[i].(float64)
		case *string:
			*v = currow.cols[i].(string)
		case *time.Time:
			*v = currow.cols[i].(time.Time)
		default:
			return ErrInvalidType
		}
	}
	return nil
}

// ErrInvalidType is returned by Scan for invalid arguments.
var ErrInvalidType = errors.New("xls: Scan only supports *int, *float64, *string, *time.Time arguments")
