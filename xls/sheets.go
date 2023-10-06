package xls

import (
	"encoding/binary"
	"errors"
	"log"
	"math"
	"unicode/utf16"

	"github.com/pbnjay/grate"
	"github.com/pbnjay/grate/commonxl"
)

// List (visible) sheet names from the workbook.
func (b *WorkBook) List() ([]string, error) {
	res := make([]string, 0, len(b.sheets))
	for _, s := range b.sheets {
		if (s.HiddenState & 0x03) == 0 {
			res = append(res, s.Name)
		}
	}
	return res, nil
}

// ListHidden sheet names in the workbook.
func (b *WorkBook) ListHidden() ([]string, error) {
	res := make([]string, 0, len(b.sheets))
	for _, s := range b.sheets {
		if (s.HiddenState & 0x03) != 0 {
			res = append(res, s.Name)
		}
	}
	return res, nil
}

// Get opens the named worksheet and return an iterator for its contents.
func (b *WorkBook) Get(sheetName string) (grate.Collection, error) {
	for _, s := range b.sheets {
		if s.Name == sheetName {
			ss := b.pos2substream[int64(s.Position)]
			return b.parseSheet(s, ss)
		}
	}
	return nil, errors.New("xls: sheet not found")
}

func (b *WorkBook) parseSheet(s *boundSheet, ss int) (*commonxl.Sheet, error) {
	res := &commonxl.Sheet{
		Formatter: &b.nfmt,
	}
	var minRow, maxRow uint32
	var minCol, maxCol uint16

	// temporary string buffer
	us := make([]uint16, 8224)

	inSubstream := 0
	for idx, r := range b.substreams[ss] {
		if inSubstream > 0 {
			if r.RecType == RecTypeEOF {
				inSubstream--
			}
			continue
		}
		switch r.RecType {
		case RecTypeBOF:
			// a BOF inside a sheet usually means embedded content like a chart
			// (which we aren't interested in). So we we set a flag and wait
			// for the EOF for that content block.
			if idx > 0 {
				inSubstream++
				continue
			}
		case RecTypeWsBool:
			if (r.Data[1] & 0x10) != 0 {
				// it's a dialog
				return nil, nil
			}

		case RecTypeDimensions:
			// max = 0-based index of the row AFTER the last valid index
			minRow = binary.LittleEndian.Uint32(r.Data[:4])
			maxRow = binary.LittleEndian.Uint32(r.Data[4:8]) // max = 0x010000
			minCol = binary.LittleEndian.Uint16(r.Data[8:10])
			maxCol = binary.LittleEndian.Uint16(r.Data[10:12]) // max = 0x000100
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

			// pre-allocate cells
			res.Resize(int(maxRow), int(maxCol))
		}
	}
	inSubstream = 0

	var formulaRow, formulaCol uint16
	for ridx, r := range b.substreams[ss] {
		if inSubstream > 0 {
			if r.RecType == RecTypeEOF {
				inSubstream--
			} else if grate.Debug {
				log.Println("      Unhandled sheet substream record type:", r.RecType, ridx)
			}
			continue
		}

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

		case RecTypeBoolErr:
			rowIndex := int(binary.LittleEndian.Uint16(r.Data[:2]))
			colIndex := int(binary.LittleEndian.Uint16(r.Data[2:4]))
			ixfe := int(binary.LittleEndian.Uint16(r.Data[4:6]))
			if r.Data[7] == 0 {
				// Boolean value
				bv := false
				if r.Data[6] == 1 {
					bv = true
				}
				var fno uint16
				if ixfe < len(b.xfs) {
					fno = b.xfs[ixfe]
				}
				res.Put(rowIndex, colIndex, bv, fno)
				//log.Printf("bool/error spec: %d %d %+v", rowIndex, colIndex, bv)
			} else {
				// it's an error, load the label
				be, ok := berrLookup[r.Data[6]]
				if !ok {
					be = "<unknown error>"
				}
				res.Put(rowIndex, colIndex, be, 0)
				//log.Printf("bool/error spec: %d %d %s", rowIndex, colIndex, be)
			}

		case RecTypeMulRk:
			// MulRk encodes multiple RK values in a row
			nrk := int((r.RecSize - 6) / 6)
			rowIndex := int(binary.LittleEndian.Uint16(r.Data[:2]))
			colIndex := int(binary.LittleEndian.Uint16(r.Data[2:4]))
			for i := 0; i < nrk; i++ {
				off := 4 + i*6
				ixfe := int(binary.LittleEndian.Uint16(r.Data[off:]))
				value := RKNumber(binary.LittleEndian.Uint32(r.Data[off+2:]))

				var rval interface{}
				if value.IsInteger() {
					rval = value.Int()
				} else {
					rval = value.Float64()
				}
				var fno uint16
				if ixfe < len(b.xfs) {
					fno = b.xfs[ixfe]
				}
				res.Put(rowIndex, colIndex+i, rval, fno)
			}
			//log.Printf("mulrow spec: %+v", *mr)

		case RecTypeNumber:
			rowIndex := int(binary.LittleEndian.Uint16(r.Data[:2]))
			colIndex := int(binary.LittleEndian.Uint16(r.Data[2:4]))
			ixfe := int(binary.LittleEndian.Uint16(r.Data[4:6]))
			xnum := binary.LittleEndian.Uint64(r.Data[6:])

			value := math.Float64frombits(xnum)
			var fno uint16
			if ixfe < len(b.xfs) {
				fno = b.xfs[ixfe]
			}
			res.Put(rowIndex, colIndex, value, fno)
			//log.Printf("Number spec: %d %d = %f", rowIndex, colIndex, value)

		case RecTypeRK:
			rowIndex := int(binary.LittleEndian.Uint16(r.Data[:2]))
			colIndex := int(binary.LittleEndian.Uint16(r.Data[2:4]))
			ixfe := int(binary.LittleEndian.Uint16(r.Data[4:]))
			value := RKNumber(binary.LittleEndian.Uint32(r.Data[6:]))

			var rval interface{}
			if value.IsInteger() {
				rval = value.Int()
			} else {
				rval = value.Float64()
			}
			var fno uint16
			if ixfe < len(b.xfs) {
				fno = b.xfs[ixfe]
			}
			res.Put(rowIndex, colIndex, rval, fno)
			//log.Printf("RK spec: %d %d = %+v", rowIndex, colIndex, rval)

		case RecTypeFormula:
			formulaRow = binary.LittleEndian.Uint16(r.Data[:2])
			formulaCol = binary.LittleEndian.Uint16(r.Data[2:4])
			ixfe := int(binary.LittleEndian.Uint16(r.Data[4:6]))
			fdata := r.Data[6:]
			var fno uint16
			if ixfe < len(b.xfs) {
				fno = b.xfs[ixfe]
			}
			if fdata[6] == 0xFF && fdata[7] == 0xFF {
				switch fdata[0] {
				case 0:
					// string in next record
					// put placeholder now to record the numFmt
					res.Put(int(formulaRow), int(formulaCol), "", fno)
				case 1:
					// boolean
					bv := false
					if fdata[2] != 0 {
						bv = true
					}
					res.Put(int(formulaRow), int(formulaCol), bv, fno)
				case 2:
					// error value
					be, ok := berrLookup[fdata[2]]
					if !ok {
						be = "<unknown error>"
					}
					res.Put(int(formulaRow), int(formulaCol), be, 0)
				case 3:
					// blank string
				default:
					log.Printf("unknown formula value type %d", fdata[0])
				}
			} else {
				xnum := binary.LittleEndian.Uint64(fdata)
				value := math.Float64frombits(xnum)
				res.Put(int(formulaRow), int(formulaCol), value, fno)
			}
			//log.Printf("formula spec: %d %d ~~ %+v", formulaRow, formulaCol, r.Data)

		case RecTypeString:
			// String is the previously rendered value of a formula
			// NB similar to the workbook SST, this can continue over
			// addition records up to 32k characters. A 1-byte flag
			// at each gap indicates if the encoding switches
			// to/from 8/16-bit characters.

			charCount := binary.LittleEndian.Uint16(r.Data[:2])
			flags := r.Data[2]
			fstr := ""
			if (flags & 1) == 0 {
				fstr = string(r.Data[3:])
			} else {
				raw := r.Data[3:]
				if int(charCount) > cap(us) {
					us = make([]uint16, charCount)
				}
				us = us[:charCount]
				for i := 0; i < int(charCount); i++ {
					us[i] = binary.LittleEndian.Uint16(raw)
					raw = raw[2:]
				}
				fstr = string(utf16.Decode(us))
			}

			if (ridx + 1) < len(b.substreams[ss]) {
				ridx2 := ridx + 1
				nrecs := len(b.substreams[ss])
				for ridx2 < nrecs {
					r2 := b.substreams[ss][ridx2]
					if r2.RecType != RecTypeContinue {
						break
					}
					if (r2.Data[0] & 1) == 0 {
						fstr += string(r2.Data[1:])
					} else {
						raw := r2.Data[1:]
						slen := len(raw) / 2
						us = us[:slen]
						for i := 0; i < slen; i++ {
							us[i] = binary.LittleEndian.Uint16(raw)
							raw = raw[2:]
						}
						fstr += string(utf16.Decode(us))
					}
					ridx2++
				}
			}
			res.Set(int(formulaRow), int(formulaCol), fstr)
			//log.Printf("String direct: %d %d '%s'", int(formulaRow), int(formulaCol), fstr)

		case RecTypeLabelSst:
			rowIndex := int(binary.LittleEndian.Uint16(r.Data[:2]))
			colIndex := int(binary.LittleEndian.Uint16(r.Data[2:4]))
			ixfe := int(binary.LittleEndian.Uint16(r.Data[4:6]))
			sstIndex := int(binary.LittleEndian.Uint32(r.Data[6:]))
			if sstIndex > len(b.strings) {
				return nil, errors.New("xls: invalid sst index")
			}
			var fno uint16
			if ixfe < len(b.xfs) {
				fno = b.xfs[ixfe]
			}
			if b.strings[sstIndex] != "" {
				res.Put(rowIndex, colIndex, b.strings[sstIndex], fno)
			}
			//log.Printf("SST spec: %d %d = [%d] '%s' %d", rowIndex, colIndex, sstIndex, b.strings[sstIndex], fno)

		case RecTypeHLink:
			firstRow := binary.LittleEndian.Uint16(r.Data[:2])
			lastRow := binary.LittleEndian.Uint16(r.Data[2:4])
			firstCol := binary.LittleEndian.Uint16(r.Data[4:6])
			lastCol := binary.LittleEndian.Uint16(r.Data[6:])
			if int(firstCol) > int(maxCol) {
				//log.Println("invalid hyperlink column")
				continue
			}
			if int(firstRow) > int(maxRow) {
				//log.Println("invalid hyperlink row")
				continue
			}
			if lastRow == 0xFFFF { // placeholder value indicate "last"
				lastRow = uint16(maxRow) - 1
			}
			if lastCol == 0xFF { // placeholder value indicate "last"
				lastCol = uint16(maxCol) - 1
			}

			// decode the hyperlink datastructure and try to find the
			// display text and separate the URL itself.
			displayText, linkText, err := decodeHyperlinks(r.Data[8:])
			if err != nil {
				log.Println(err)
				continue
			}

			// apply merge cell rules (see RecTypeMergeCells below)
			for rn := int(firstRow); rn <= int(lastRow); rn++ {
				for cn := int(firstCol); cn <= int(lastCol); cn++ {
					if rn == int(firstRow) && cn == int(firstCol) {
						// TODO: provide custom hooks for how to handle links in output
						res.Put(rn, cn, displayText+" <"+linkText+">", 0)
					} else if cn == int(firstCol) {
						// first and last column MAY be the same
						if rn == int(lastRow) {
							res.Put(rn, cn, grate.EndRowMerged, 0)
						} else {
							res.Put(rn, cn, grate.ContinueRowMerged, 0)
						}
					} else if cn == int(lastCol) {
						// first and last column are NOT the same
						res.Put(rn, cn, grate.EndColumnMerged, 0)
					} else {
						res.Put(rn, cn, grate.ContinueColumnMerged, 0)
					}
				}
			}

		case RecTypeMergeCells:
			// To keep cells aligned, Merged cells are handled by placing
			// special characters in each cell covered by the merge block.
			//
			// The contents of the cell are always in the top left position.
			// A "down arrow" (↓) indicates the left side of the merge block, and a
			// "down arrow with stop line" (⤓) indicates the last row of the merge.
			// A "right arrow" (→) indicates that the columns span horizontally,
			// and a "right arrow with stop line" (⇥) indicates the rightmost
			// column of the merge.
			//

			cmcs := binary.LittleEndian.Uint16(r.Data[:2])
			raw := r.Data[2:]
			for i := 0; i < int(cmcs); i++ {
				firstRow := binary.LittleEndian.Uint16(raw[:2])
				lastRow := binary.LittleEndian.Uint16(raw[2:4])
				firstCol := binary.LittleEndian.Uint16(raw[4:6])
				lastCol := binary.LittleEndian.Uint16(raw[6:])
				raw = raw[8:]

				if lastRow == 0xFFFF { // placeholder value indicate "last"
					lastRow = uint16(maxRow) - 1
				}
				if lastCol == 0xFF { // placeholder value indicate "last"
					lastCol = uint16(maxCol) - 1
				}
				for rn := int(firstRow); rn <= int(lastRow); rn++ {
					for cn := int(firstCol); cn <= int(lastCol); cn++ {
						if rn == int(firstRow) && cn == int(firstCol) {
							// should be a value there already!
						} else if cn == int(firstCol) {
							// first and last column MAY be the same
							if rn == int(lastRow) {
								res.Put(rn, cn, grate.EndRowMerged, 0)
							} else {
								res.Put(rn, cn, grate.ContinueRowMerged, 0)
							}
						} else if cn == int(lastCol) {
							// first and last column are NOT the same
							res.Put(rn, cn, grate.EndColumnMerged, 0)
						} else {
							res.Put(rn, cn, grate.ContinueColumnMerged, 0)
						}
					}
				}
			}
			/*
				case RecTypeBlank, RecTypeMulBlank:
					// cells default value is blank, no need for these

				case RecTypeContinue:
					// the only situation so far is when used in RecTypeString above

				case RecTypeRow, RecTypeDimensions, RecTypeEOF, RecTypeWsBool:
					// handled in initial pass

				default:
					if grate.Debug {
						log.Println("    Unhandled sheet record type:", r.RecType, ridx)
					}
			*/
		}
	}
	return res, nil
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
