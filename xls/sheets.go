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

	"github.com/pbnjay/grate"
)

func (b *WorkBook) List() ([]string, error) {
	res := make([]string, 0, len(b.sheets))
	for _, s := range b.sheets {
		if (s.HiddenState & 0x03) == 0 {
			res = append(res, s.Name)
		}
	}
	return res, nil
}

func (b *WorkBook) Get(sheetName string) (grate.Collection, error) {
	for _, s := range b.sheets {
		if s.Name == sheetName {
			ss := b.pos2substream[int64(s.Position)]
			ws := &WorkSheet{
				b: b, s: s, ss: ss,
				iterRow: -1,
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

	minRow int
	maxRow int // maximum valid row index (0xFFFF)
	minCol int
	maxCol int // maximum valid column index (0xFF)
	rows   []*row
	empty  bool

	iterRow int
	iterMC  int
}

type staticCellType rune

const (
	staticBlank staticCellType = 0

	// marks a continuation column within a merged cell.
	continueColumnMerged staticCellType = '→'
	// marks the last column of a merged cell.
	endColumnMerged staticCellType = '⇥'

	// marks a continuation row within a merged cell.
	continueRowMerged staticCellType = '↓'
	// marks the last row of a merged cell.
	endRowMerged staticCellType = '⤓'
)

func (s staticCellType) String() string {
	if s == 0 {
		return ""
	}
	return string([]rune{rune(s)})
}

type row struct {
	// each value must be one of: int, float64, string, or time.Time
	cols []interface{}
}

func (s *WorkSheet) placeValue(rowIndex, colIndex int, val interface{}) {
	if colIndex > s.maxCol || rowIndex > s.maxRow {
		// invalid
		return
	}

	// ensure we always have a complete matrix
	for len(s.rows) <= rowIndex {
		emptyRow := make([]interface{}, s.maxCol+1)
		for i := 0; i <= s.maxCol; i++ {
			emptyRow[i] = staticBlank
		}
		s.rows = append(s.rows, &row{emptyRow})
	}

	s.rows[rowIndex].cols[colIndex] = val
}

func (s *WorkSheet) IsEmpty() bool {
	return s.empty
}

func (s *WorkSheet) parse() error {
	inSubstream := 0
	for idx, r := range s.b.substreams[s.ss] {
		if inSubstream > 0 {
			if r.RecType == RecTypeEOF {
				inSubstream--
			}
			continue
		}
		switch r.RecType {
		case RecTypeBOF:
			if idx > 0 {
				inSubstream++
				continue
			}
		case RecTypeWsBool:
			if (r.Data[1] & 0x10) != 0 {
				// it's a dialog
				return nil
			}

		case RecTypeDimensions:
			bb := bytes.NewReader(r.Data)
			var minRow, maxRow uint32
			var minCol, maxCol uint16

			// max = 0-based index of the row AFTER the last valid index
			binary.Read(bb, binary.LittleEndian, &minRow)
			binary.Read(bb, binary.LittleEndian, &maxRow) // max = 0x010000
			binary.Read(bb, binary.LittleEndian, &minCol)
			binary.Read(bb, binary.LittleEndian, &maxCol) // max = 0x000100
			if grate.Debug {
				log.Printf("    Sheet dimensions (%d, %d) - (%d,%d)",
					minCol, minRow, maxCol, maxRow)
			}
			if minRow > 0x0000FFFF || maxRow > 0x00010000 {
				log.Println("invalid dimensions")
			}
			if minCol > 0x00FF || maxCol > 0x0100 {
				log.Println("invalid dimensions")
			}
			s.minRow = int(uint64(minRow) & 0x0FFFF)
			s.maxRow = int(uint64(maxRow)&0x1FFFF) - 1 // translate to last valid index
			s.minCol = int(uint64(minCol) & 0x000FF)
			s.maxCol = int(uint64(maxCol)&0x001FF) - 1 // translate to last valid index
			if (maxRow-minRow) == 0 || (maxCol-minCol) == 0 {
				s.empty = true
			} else {
				// pre-allocate cells
				s.placeValue(s.maxRow, s.maxCol, staticBlank)
			}

		case RecTypeRow:
			bb := bytes.NewReader(r.Data)
			row := &shRow{}
			binary.Read(bb, binary.LittleEndian, row)
			if (row.Reserved & 0xFFFF) != 0 {
				log.Println("invalid Row spec")
				continue
			}
			//log.Printf("row spec: %+v", *row)
		}
	}
	inSubstream = 0

	var formulaRow, formulaCol uint16
	for ridx, r := range s.b.substreams[s.ss] {
		if inSubstream > 0 {
			if r.RecType == RecTypeEOF {
				inSubstream--
			} else if grate.Debug {
				log.Println("      Unhandled sheet substream record type:", r.RecType, ridx)
			}
			continue
		}

		bb := bytes.NewReader(r.Data)

		// sec 2.1.7.20.6 Common Productions ABNF:
		/*
			CELLTABLE = 1*(1*Row *CELL 1*DBCell) *EntExU2
			CELL = FORMULA / Blank / MulBlank / RK / MulRk / BoolErr / Number / LabelSst
			FORMULA = [Uncalced] Formula [Array / Table / ShrFmla / SUB] [String *Continue]

			Not parsed form the list above:
				DBCell, EntExU2, Uncalced, Array, Table,ShrFmla
				NB: no idea what "SUB" is
		*/

		switch r.RecType {
		case RecTypeBOF:
			if ridx > 0 {
				inSubstream++
				continue
			}
		case RecTypeBlank:
			var rowIndex, colIndex uint16
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &colIndex)
			//log.Printf("blank spec: %d %d", rowIndex, colIndex)

		case RecTypeBoolErr:
			var rowIndex, colIndex, ixfe uint16
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &colIndex)
			binary.Read(bb, binary.LittleEndian, &ixfe)
			if r.Data[7] == 0 {
				bv := false
				if r.Data[6] == 1 {
					bv = true
				}
				s.placeValue(int(rowIndex), int(colIndex), bv)
				//log.Printf("bool/error spec: %d %d %+v", rowIndex, colIndex, bv)
			} else {
				be, ok := berrLookup[r.Data[6]]
				if !ok {
					be = "<unknown error>"
				}
				s.placeValue(int(rowIndex), int(colIndex), be)
				//log.Printf("bool/error spec: %d %d %s", rowIndex, colIndex, be)
			}

		case RecTypeMulBlank:
			var rowIndex, firstCol uint16
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &firstCol)
		//	nrk := int((r.RecSize - 6) / 6)
		//	log.Printf("row blanks spec: %d %d %d", rowIndex, firstCol, nrk)

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
					fno := s.b.xfs[rr.IXFCell]
					rval, _ = s.b.nfmt.Apply(fno, rval)
				}
				s.placeValue(int(mr.RowIndex), int(mr.FirstCol)+i, rval)
			}
			binary.Read(bb, binary.LittleEndian, &mr.LastCol)
			//log.Printf("mulrow spec: %+v", *mr)

		case RecTypeNumber:
			var rowIndex, colIndex, ixfe uint16
			var xnum uint64
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &colIndex)
			binary.Read(bb, binary.LittleEndian, &ixfe)
			binary.Read(bb, binary.LittleEndian, &xnum)
			value := math.Float64frombits(xnum)
			fno := s.b.xfs[ixfe]
			rval, _ := s.b.nfmt.Apply(fno, value)

			s.placeValue(int(rowIndex), int(colIndex), rval)
			//log.Printf("Number spec: %d %d = %f", rowIndex, colIndex, value)

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
				fno := s.b.xfs[rr.IXFCell]
				rval, _ = s.b.nfmt.Apply(fno, rval)
			}
			s.placeValue(int(rowIndex), int(colIndex), rval)
			//log.Printf("RK spec: %d %d = %s", rowIndex, colIndex, rr.Value.String())

		case RecTypeFormula:
			var ixfe uint16
			binary.Read(bb, binary.LittleEndian, &formulaRow)
			binary.Read(bb, binary.LittleEndian, &formulaCol)
			binary.Read(bb, binary.LittleEndian, &ixfe)
			fdata := r.Data[6:]
			if fdata[6] == 0xFF && r.Data[7] == 0xFF {
				switch fdata[0] {
				case 0:
					// string in next record
				case 1:
					// boolean
					bv := false
					if fdata[2] != 0 {
						bv = true
					}
					s.placeValue(int(formulaRow), int(formulaCol), bv)
				case 2:
					// error value
					be, ok := berrLookup[fdata[2]]
					if !ok {
						be = "<unknown error>"
					}
					s.placeValue(int(formulaRow), int(formulaCol), be)
				case 3:
					// blank string
				default:
					log.Println("unknown formula value type")
				}
			} else {
				var xnum uint64
				binary.Read(bb, binary.LittleEndian, &xnum)
				value := math.Float64frombits(xnum)
				fno := s.b.xfs[ixfe]
				rval, _ := s.b.nfmt.Apply(fno, value)
				s.placeValue(int(formulaRow), int(formulaCol), rval)
			}
			//log.Printf("formula spec: %d %d ~~ %+v", formulaRow, formulaCol, r.Data)

		case RecTypeString:
			var charCount uint16
			var flags byte
			binary.Read(bb, binary.LittleEndian, &charCount)
			binary.Read(bb, binary.LittleEndian, &flags)
			fstr := ""
			if (flags & 1) == 0 {
				fstr = string(r.Data[3:])
			} else {
				us := make([]uint16, charCount)
				binary.Read(bb, binary.LittleEndian, us)
				fstr = string(utf16.Decode(us))
			}

			if (ridx + 1) < len(s.b.substreams[s.ss]) {
				ridx2 := ridx + 1
				nrecs := len(s.b.substreams[s.ss])
				for ridx2 < nrecs {
					r2 := s.b.substreams[s.ss][ridx2]
					if r2.RecType != RecTypeContinue {
						break
					}
					if (r2.Data[0] & 1) == 0 {
						fstr += string(r2.Data[1:])
					} else {
						bb2 := bytes.NewReader(r2.Data[1:])
						us := make([]uint16, len(r2.Data)-1)
						binary.Read(bb2, binary.LittleEndian, us)
						fstr += string(utf16.Decode(us))
					}
					ridx2++
				}
			}
			// TODO: does formula record formatted dates as pre-computed strings?
			s.placeValue(int(formulaRow), int(formulaCol), fstr)

		case RecTypeLabelSst:
			var rowIndex, colIndex, ixfe uint16
			var sstIndex uint32
			binary.Read(bb, binary.LittleEndian, &rowIndex)
			binary.Read(bb, binary.LittleEndian, &colIndex)
			binary.Read(bb, binary.LittleEndian, &ixfe)
			binary.Read(bb, binary.LittleEndian, &sstIndex)
			if int(sstIndex) > len(s.b.strings) {
				return errors.New("xls: invalid sst index")
			}
			s.placeValue(int(rowIndex), int(colIndex), s.b.strings[sstIndex])
			//log.Printf("SST spec: %d %d = [%d] %s", rowIndex, colIndex, sstIndex, s.b.strings[sstIndex])

		case RecTypeHLink:
			loc := &shRef8{}
			binary.Read(bb, binary.LittleEndian, loc)
			if int(loc.FirstCol) > s.maxCol {
				//log.Println("invalid hyperlink column")
				continue
			}
			if int(loc.FirstRow) > s.maxRow {
				//log.Println("invalid hyperlink row")
				continue
			}
			if loc.LastRow == 0xFFFF {
				loc.LastRow = uint16(s.maxRow)
			}
			if loc.LastCol == 0xFF {
				loc.LastCol = uint16(s.maxCol)
			}

			displayText, linkText, err := decodeHyperlinks(bb)
			if err != nil {
				log.Println(err)
				continue
			}

			// apply merge cell rules
			for rn := int(loc.FirstRow); rn <= int(loc.LastRow); rn++ {
				for cn := int(loc.FirstCol); cn <= int(loc.LastCol); cn++ {
					if rn == int(loc.FirstRow) && cn == int(loc.FirstCol) {
						s.placeValue(rn, cn, displayText+" <"+linkText+">")
					} else if cn == int(loc.FirstCol) {
						// first and last column MAY be the same
						if rn == int(loc.LastRow) {
							s.placeValue(rn, cn, endRowMerged)
						} else {
							s.placeValue(rn, cn, continueRowMerged)
						}
					} else if cn == int(loc.LastCol) {
						// first and last column are NOT the same
						s.placeValue(rn, cn, endColumnMerged)
					} else {
						s.placeValue(rn, cn, continueColumnMerged)
					}
				}
			}

		case RecTypeMergeCells:
			var cmcs uint16
			binary.Read(bb, binary.LittleEndian, &cmcs)
			mcRefs := make([]shRef8, cmcs)
			binary.Read(bb, binary.LittleEndian, &mcRefs)
			for _, loc := range mcRefs {
				if loc.LastRow == 0xFFFF {
					loc.LastRow = uint16(s.maxRow)
				}
				if loc.LastCol == 0xFF {
					loc.LastCol = uint16(s.maxCol)
				}
				for rn := int(loc.FirstRow); rn <= int(loc.LastRow); rn++ {
					for cn := int(loc.FirstCol); cn <= int(loc.LastCol); cn++ {
						if rn == int(loc.FirstRow) && cn == int(loc.FirstCol) {
							// should be a value there already!
						} else if cn == int(loc.FirstCol) {
							// first and last column MAY be the same
							if rn == int(loc.LastRow) {
								s.placeValue(rn, cn, endRowMerged)
							} else {
								s.placeValue(rn, cn, continueRowMerged)
							}
						} else if cn == int(loc.LastCol) {
							// first and last column are NOT the same
							s.placeValue(rn, cn, endColumnMerged)
						} else {
							s.placeValue(rn, cn, continueColumnMerged)
						}
					}
				}
			}

		case RecTypeContinue:
			// the only situation so far is when used in RecTypeString above

		case RecTypeRow, RecTypeDimensions, RecTypeEOF, RecTypeWsBool:
			// handled in initial pass
		default:
			if grate.Debug {
				log.Println("    Unhandled sheet record type:", r.RecType, ridx)
			}
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
// Arguments must be pointers to one of 5 supported types:
//     bool, int, float64, string, or time.Time
func (s *WorkSheet) Scan(args ...interface{}) error {
	currow := s.rows[s.iterRow]

	for i, a := range args {
		switch v := a.(type) {
		case *bool:
			*v = currow.cols[i].(bool)
		case *int:
			*v = currow.cols[i].(int)
		case *float64:
			*v = currow.cols[i].(float64)
		case *string:
			*v = currow.cols[i].(string)
		case *time.Time:
			*v = currow.cols[i].(time.Time)
		default:
			return grate.ErrInvalidScanType
		}
	}
	return nil
}

var berrLookup = map[byte]string{
	0x00: "#NULL!",
	0x07: "#DIV/0!",
	0x0F: "#VALUE!",
	0x17: "#REF!",
	0x1D: "#NAME?",
	0x24: "#NUM!",
	0x2A: "#N/A",
	0x2B: "#GETTING_DATA",
}
