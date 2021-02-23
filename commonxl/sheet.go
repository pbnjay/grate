package commonxl

import (
	"fmt"
	"log"
	"time"
)

// Sheet holds raw and rendered values for a spreadsheet.
type Sheet struct {
	Formatter *Formatter
	NumRows   int
	NumCols   int
	Rows      [][]Cell

	CurRow int
}

// Resize the sheet for the number of rows and cols given.
// Newly added cells default to blank.
func (s *Sheet) Resize(rows, cols int) {
	// some sheets are off by one
	rows++
	cols++

	if rows <= 0 {
		rows = 1
	}
	if cols <= 0 {
		cols = 1
	}
	s.CurRow = 0
	s.NumRows = rows
	s.NumCols = cols

	for rows >= len(s.Rows) {
		s.Rows = append(s.Rows, make([]Cell, cols))
	}

	for i := 0; len(s.Rows[i]) < cols; i++ {
		r2 := make([]Cell, cols-len(s.Rows[i]))
		s.Rows[i] = append(s.Rows[i], r2...)
	}
}

// Put the value at the cell location given.
func (s *Sheet) Put(row, col int, value interface{}, fmtNum uint16) {
	if row >= s.NumRows || col >= s.NumCols {
		log.Printf("grate: cell out of bounds row %d>=%d, col %d>=%d",
			row, s.NumRows, col, s.NumCols)
		return
	}

	ct, ok := s.Formatter.getCellType(fmtNum)
	if !ok || fmtNum == 0 {
		s.Rows[row][col] = NewCell(value)
	} else {
		s.Rows[row][col] = NewCellWithType(value, ct, s.Formatter)
	}
	s.Rows[row][col].SetFormatNumber(fmtNum)
}

// Set changes the value in an existing cell location.
// NB Currently only used for populating string results for formulas.
func (s *Sheet) Set(row, col int, value interface{}) {
	if row > s.NumRows || col > s.NumCols {
		log.Println("grate: cell out of bounds")
		return
	}

	s.Rows[row][col][0] = value
	s.Rows[row][col][1] = StringCell
}

// SetURL adds a hyperlink to an existing cell location.
func (s *Sheet) SetURL(row, col int, link string) {
	if row > s.NumRows || col > s.NumCols {
		log.Println("grate: cell out of bounds")
		return
	}

	s.Rows[row][col].SetURL(link)
}

// Next advances to the next record of content.
// It MUST be called prior to any Scan().
func (s *Sheet) Next() bool {
	if (s.CurRow + 1) > len(s.Rows) {
		return false
	}
	s.CurRow++
	return true
}

// Strings extracts values from the current record into a list of strings.
func (s *Sheet) Strings() []string {
	res := make([]string, s.NumCols)
	for i, cell := range s.Rows[s.CurRow-1] {
		if cell.Type() == BlankCell {
			res[i] = ""
			continue
		}
		val := cell.Value()
		fs, ok := s.Formatter.Apply(cell.FormatNo(), val)
		if !ok {
			fs = fmt.Sprint(val)
		}
		res[i] = fs
	}
	return res
}

// Types extracts the data types from the current record into a list.
// options: "boolean", "integer", "float", "string", "date",
// and special cases: "blank", "hyperlink" which are string types
func (s *Sheet) Types() []string {
	res := make([]string, s.NumCols)
	for i, cell := range s.Rows[s.CurRow-1] {
		res[i] = cell.Type().String()
	}
	return res
}

// Scan extracts values from the current record into the provided arguments
// Arguments must be pointers to one of 5 supported types:
//     bool, int64, float64, string, or time.Time
// If invalid, returns ErrInvalidScanType
func (s *Sheet) Scan(args ...interface{}) error {
	row := s.Rows[s.CurRow-1]

	for i, a := range args {
		val := row[i].Value()

		switch v := a.(type) {
		case bool, int64, float64, string, time.Time:
			return fmt.Errorf("scan destinations must be pointer (arg %d is not)", i)
		case *bool:
			if x, ok := val.(bool); ok {
				*v = x
			} else {
				return fmt.Errorf("scan destination %d expected *%T, not *bool", i, val)
			}
		case *int64:
			if x, ok := val.(int64); ok {
				*v = x
			} else {
				return fmt.Errorf("scan destination %d expected *%T, not *int64", i, val)
			}
		case *float64:
			if x, ok := val.(float64); ok {
				*v = x
			} else {
				return fmt.Errorf("scan destination %d expected *%T, not *float64", i, val)
			}
		case *string:
			if x, ok := val.(string); ok {
				*v = x
			} else {
				return fmt.Errorf("scan destination %d expected *%T, not *string", i, val)
			}
		case *time.Time:
			if x, ok := val.(time.Time); ok {
				*v = x
			} else {
				return fmt.Errorf("scan destination %d expected *%T, not *time.Time", i, val)
			}
		default:
			return fmt.Errorf("scan destination for arg %d is not supported (%T)", i, a)
		}
	}
	return nil
}

// IsEmpty returns true if there are no data values.
func (s *Sheet) IsEmpty() bool {
	return (s.NumCols <= 1 && s.NumRows <= 1)
}

// Err returns the last error that occured.
func (s *Sheet) Err() error {
	return nil
}
