package xlsx

import (
	"encoding/xml"
	"errors"
	"io"
	"log"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/fcwoknhenuxdfiyv/grate"
	"github.com/fcwoknhenuxdfiyv/grate/commonxl"
)

type Sheet struct {
	d       *Document
	relID   string
	name    string
	docname string

	err error

	wrapped *commonxl.Sheet
}

var errNotLoaded = errors.New("xlsx: sheet not loaded")

func (s *Sheet) parseSheet() error {
	s.wrapped = &commonxl.Sheet{
		Formatter: &s.d.fmt,
	}
	linkmap := make(map[string]string)
	base := filepath.Base(s.docname)
	sub := strings.TrimSuffix(s.docname, base)
	relsname := filepath.Join(sub, "_rels", base+".rels")
	dec, clo, err := s.d.openXML(relsname)
	if err == nil {
		// rels might not exist for every sheet
		tok, err := dec.RawToken()
		for ; err == nil; tok, err = dec.RawToken() {
			if v, ok := tok.(xml.StartElement); ok && v.Name.Local == "Relationship" {
				ax := getAttrs(v.Attr, "Id", "Type", "Target", "TargetMode")
				if ax[3] == "External" && ax[1] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" {
					linkmap[ax[0]] = ax[2]
				}
			}
		}
		clo.Close()
	}

	dec, clo, err = s.d.openXML(s.docname)
	if err != nil {
		return err
	}
	defer clo.Close()

	currentCellType := BlankCellType
	currentCell := ""
	var fno uint16
	var maxCol, maxRow int

	tok, err := dec.RawToken()
	for ; err == nil; tok, err = dec.RawToken() {
		switch v := tok.(type) {
		case xml.CharData:
			if currentCell == "" {
				continue
			}
			c, r := refToIndexes(currentCell)
			if c >= 0 && r >= 0 {
				var val interface{} = string(v)

				switch currentCellType {
				case BooleanCellType:
					if v[0] == '1' {
						val = true
					} else {
						val = false
					}
				case DateCellType:
					log.Println("CELL DATE", val, fno)
				case NumberCellType:
					fval, err := strconv.ParseFloat(string(v), 64)
					if err == nil {
						val = fval
					}
					//log.Println("CELL NUMBER", val, numFormat)
				case SharedStringCellType:
					//log.Println("CELL SHSTR", val, currentCellType, numFormat)
					si, _ := strconv.ParseInt(string(v), 10, 64)
					val = s.d.strings[si]
				case BlankCellType:
					//log.Println("CELL BLANK")
					// don't place any values
					continue
				case ErrorCellType, FormulaStringCellType, InlineStringCellType:
					//log.Println("CELL ERR/FORM/INLINE", val, currentCellType)
				default:
					log.Println("CELL UNKNOWN", val, currentCellType, fno)
				}
				s.wrapped.Put(r, c, val, fno)
			} else {
				//log.Println("FAIL row/col: ", currentCell)
			}
		case xml.StartElement:
			switch v.Name.Local {
			case "dimension":
				ax := getAttrs(v.Attr, "ref")
				if ax[0] == "A1" {
					maxCol, maxRow = 1, 1
					// short-circuit empty sheet
					s.wrapped.Resize(1, 1)
					continue
				}
				dims := strings.Split(ax[0], ":")
				if len(dims) == 1 {
					maxCol, maxRow = refToIndexes(dims[0])
				} else {
					//minCol, minRow := refToIndexes(dims[0])
					maxCol, maxRow = refToIndexes(dims[1])
				}
				s.wrapped.Resize(maxRow, maxCol)
				//log.Println("DIMENSION:", s.minRow, s.minCol, ">", s.maxRow, s.maxCol)
			case "row":
				//currentRow = ax["r"] // unsigned int row index
				//log.Println("ROW", currentRow)
			case "c":
				ax := getAttrs(v.Attr, "t", "r", "s")
				currentCellType = CellType(ax[0])
				if currentCellType == BlankCellType {
					currentCellType = NumberCellType
				}
				currentCell = ax[1] // always an A1 style reference
				style := ax[2]
				sid, _ := strconv.ParseInt(style, 10, 64)
				if len(s.d.xfs) > int(sid) {
					fno = s.d.xfs[sid]
				} else {
					fno = 0
				}
				//log.Println("CELL", currentCell, sid, numFormat, currentCellType)
			case "v":
				//log.Println("CELL VALUE", ax)

			case "mergeCell":
				ax := getAttrs(v.Attr, "ref")
				dims := strings.Split(ax[0], ":")
				startCol, startRow := refToIndexes(dims[0])
				endCol, endRow := startCol, startRow
				if len(dims) > 1 {
					endCol, endRow = refToIndexes(dims[1])
				}
				if endRow > maxRow {
					endRow = maxRow
				}
				if endCol > maxCol {
					endCol = maxCol
				}
				for r := startRow; r <= endRow; r++ {
					for c := startCol; c <= endCol; c++ {
						if r == startRow && c == startCol {
							// has data already!
						} else if c == startCol {
							// first and last column MAY be the same
							if r == endRow {
								s.wrapped.Put(r, c, grate.EndRowMerged, 0)
							} else {
								s.wrapped.Put(r, c, grate.ContinueRowMerged, 0)
							}
						} else if c == endCol {
							// first and last column are NOT the same
							s.wrapped.Put(r, c, grate.EndColumnMerged, 0)
						} else {
							s.wrapped.Put(r, c, grate.ContinueColumnMerged, 0)
						}
					}
				}

			case "hyperlink":
				ax := getAttrs(v.Attr, "ref", "id")
				col, row := refToIndexes(ax[0])
				link := linkmap[ax[1]]
				s.wrapped.Put(row, col, link, 0)
				s.wrapped.SetURL(row, col, link)

			case "worksheet", "mergeCells", "hyperlinks":
				// containers
			case "f":
				//log.Println("start: ", v.Name.Local, v.Attr)
			default:
				if grate.Debug {
					log.Println("      Unhandled sheet xml tag", v.Name.Local, v.Attr)
				}
			}
		case xml.EndElement:

			switch v.Name.Local {
			case "c":
				currentCell = ""
			case "row":
				//currentRow = ""
			}
		default:
			if grate.Debug {
				log.Printf("      Unhandled sheet xml tokens %T %+v", tok, tok)
			}
		}
	}
	if err == io.EOF {
		err = nil
	}
	return err
}
