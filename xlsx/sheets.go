package xlsx

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"log"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

type Sheet struct {
	d       *Document
	relID   string
	name    string
	docname string

	err error

	minRow int
	maxRow int
	minCol int
	maxCol int
	rows   []*row
	empty  bool

	iterRow int
}

type row struct {
	// each value must be one of: int, float64, string, or time.Time
	cols []interface{}
}

func (s *Sheet) parseSheet() error {
	linkmap := make(map[string]string)
	base := filepath.Base(s.docname)
	sub := strings.TrimSuffix(s.docname, base)
	relsname := filepath.Join(sub, "_rels", base+".rels")
	dec, clo, err := s.d.openXML(relsname)
	if err == nil {
		// rels might not exist for every sheet
		tok, err := dec.Token()
		for ; err == nil; tok, err = dec.Token() {
			if v, ok := tok.(xml.StartElement); ok && v.Name.Local == "Relationship" {
				ax := attrMap(v.Attr)
				if ax["TargetMode"] == "External" && ax["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" {
					linkmap[ax["Id"]] = ax["Target"]
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
	numFormat := ""
	tok, err := dec.Token()
	for ; err == nil; tok, err = dec.Token() {
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
					log.Println("CELL DATE", val, numFormat)
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
					log.Println("CELL UNKNOWN", val, currentCellType, numFormat)
				}
				s.placeValue(r, c, val)
			} else {
				//log.Println("FAIL row/col: ", currentCell)
			}
		case xml.StartElement:
			ax := attrMap(v.Attr)
			switch v.Name.Local {
			case "dimension":
				if ax["ref"] == "A1" {
					// short-circuit empty sheet
					s.minCol, s.minRow = 0, 0
					s.maxCol, s.maxRow = 1, 1
					s.empty = true
					continue
				}
				dims := strings.Split(ax["ref"], ":")
				s.minCol, s.minRow = refToIndexes(dims[0])
				s.maxCol, s.maxRow = refToIndexes(dims[1])
				//log.Println("DIMENSION:", s.minRow, s.minCol, ">", s.maxRow, s.maxCol)
			case "row":
				//currentRow = ax["r"] // unsigned int row index
				//log.Println("ROW", currentRow)
			case "c":
				currentCellType = CellType(ax["t"])
				if currentCellType == BlankCellType {
					currentCellType = NumberCellType
				}
				currentCell = ax["r"] // always an A1 style reference
				style := ax["s"]
				sid, _ := strconv.ParseInt(style, 10, 64)
				numFormat = s.d.xfs[sid] // unsigned integer lookup
				//log.Println("CELL", currentCell, sid, numFormat, currentCellType)
			case "v":
				//log.Println("CELL VALUE", ax)

			case "mergeCell":
				dims := strings.Split(ax["ref"], ":")
				startCol, startRow := refToIndexes(dims[0])
				endCol, endRow := refToIndexes(dims[1])
				for r := startRow; r <= endRow; r++ {
					for c := startCol; c <= endCol; c++ {
						if r == startRow && c == startCol {
							// has data already!
						} else if c == startCol {
							// first and last column MAY be the same
							if r == endRow {
								s.placeValue(r, c, endRowMerged)
							} else {
								s.placeValue(r, c, continueRowMerged)
							}
						} else if c == endCol {
							// first and last column are NOT the same
							s.placeValue(r, c, endColumnMerged)
						} else {
							s.placeValue(r, c, continueColumnMerged)
						}
					}
				}

			case "hyperlink":
				col, row := refToIndexes(ax["ref"])
				link := linkmap[ax["id"]]
				if len(s.rows) > row && len(s.rows[row].cols) > col {
					if sstr, ok := s.rows[row].cols[col].(string); ok {
						link = sstr + " <" + link + ">"
					}
				}
				s.placeValue(row, col, link)

			case "mergeCells", "hyperlinks":
				// NB don't need these outer containers
			case "f":
				//log.Println("start: ", v.Name.Local, v.Attr)
			default:
				//log.Println("start: ", v.Name.Local, v.Attr)
			}
		case xml.EndElement:

			switch v.Name.Local {
			case "c":
				currentCell = ""
			case "row":
				//currentRow = ""
			}
			//log.Println("  end: ", v.Name.Local)
		default:
			//log.Printf("%T %+v", tok, tok)
		}
	}
	if err == io.EOF {
		err = nil
	}
	return err
}

func (s *Sheet) placeValue(rowIndex, colIndex int, val interface{}) {
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
	s.empty = false
	s.rows[rowIndex].cols[colIndex] = val
}

// Next advances to the next row of content.
// It MUST be called prior to any Scan().
func (s *Sheet) Next() bool {
	s.iterRow++
	return s.iterRow < len(s.rows)
}

func (s *Sheet) Strings() []string {
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
func (s *Sheet) Scan(args ...interface{}) error {
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
			return ErrInvalidType
		}
	}
	return nil
}

func (s *Sheet) IsEmpty() bool {
	return s.empty
}

// Err returns the last error that occured.
func (s *Sheet) Err() error {
	return s.err
}

// ErrInvalidType is returned by Scan for invalid arguments.
var ErrInvalidType = errors.New("xlsx: Scan only supports *bool, *int, *float64, *string, *time.Time arguments")
