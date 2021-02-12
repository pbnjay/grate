package xlsx

import (
	"encoding/xml"
	"strconv"
	"strings"
)

type CellType string

// CellTypes define data type in section 18.18.11
const (
	BlankCellType         CellType = ""
	BooleanCellType       CellType = "b"
	DateCellType          CellType = "d"
	ErrorCellType         CellType = "e"
	NumberCellType        CellType = "n"
	SharedStringCellType  CellType = "s"
	FormulaStringCellType CellType = "str"
	InlineStringCellType  CellType = "inlineStr"
)

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

// returns the 0-based index of the column string:
//    "A"=0, "B"=1, "AA"=26, "BB"=53
func col2int(col string) int {
	idx := 0
	for _, c := range col {
		idx *= 26
		idx += int(c - '@')
	}
	return idx - 1
}

func refToIndexes(r string) (column, row int) {
	if len(r) < 2 {
		return -1, -1
	}
	i1 := strings.IndexAny(r, "0123456789")
	if i1 <= 0 {
		return -1, -1
	}

	// A1 Reference mode
	col1 := r[:i1]
	i2 := strings.IndexByte(r[i1:], 'C')
	if i2 == -1 {
		rn, _ := strconv.ParseInt(r[i1:], 10, 64)
		return col2int(col1), int(rn)
	}

	// R1C1 Reference Mode
	col1 = r[i1:i2]
	row1 := r[i2+1:]
	cn, _ := strconv.ParseInt(col1, 10, 64)
	rn, _ := strconv.ParseInt(row1, 10, 64)
	return int(cn), int(rn)
}

func getAttrs(attrs []xml.Attr, keys ...string) []string {
	res := make([]string, len(keys))
	for _, a := range attrs {
		for i, k := range keys {
			if a.Name.Local == k {
				res[i] = a.Value
			}
		}
	}
	return res
}
