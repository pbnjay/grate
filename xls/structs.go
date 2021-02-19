package xls

import (
	"fmt"
	"math"
)

type header struct {
	Version  uint16 // An unsigned integer that specifies the BIFF version of the file. The value MUST be 0x0600.
	DocType  uint16 //An unsigned integer that specifies the document type of the substream of records following this record. For more information about the layout of the sub-streams in the workbook stream see File Structure.
	RupBuild uint16 // An unsigned integer that specifies the build identifier.
	RupYear  uint16 // An unsigned integer that specifies the year when this BIFF version was first created. The value MUST be 0x07CC or 0x07CD.
	MiscBits uint64 // lots of miscellaneous bits and flags we're not going to check
}

// 2.1.4
type rec struct {
	Data    []byte
	RecType recordType
	RecSize uint16
}

type boundSheet struct {
	Name        string
	Position    uint32
	HiddenState byte
	SheetType   byte
}

///////
type shRow struct {
	RowIndex uint16 // 0-based
	FirstCol uint16 // 0-based
	LastCol  uint16 // 1-based!
	Height   uint16
	Reserved uint32
	Flags    uint32
}

type shRef8 struct {
	FirstRow uint16 // 0-based
	LastRow  uint16 // 0-based
	FirstCol uint16 // 0-based
	LastCol  uint16 // 0-based
}
type shMulRK struct {
	Values   []RkRec
	RowIndex uint16
	FirstCol uint16
	LastCol  uint16
}
type RkRec struct {
	IXFCell uint16
	Value   RKNumber
}

type shRK struct {
	RowIndex uint16 // 0-based
	Col      uint16 // 0-based
	IXFCell  uint16
	Value    RKNumber
}

type RKNumber uint32

func (r RKNumber) IsInteger() bool {
	if (r & 1) != 0 {
		// has 2 decimals
		return false
	}
	if (r & 2) == 0 {
		// is part of a float
		return false
	}
	return true
}

func (r RKNumber) Int() int {
	val := int32(r) >> 2
	if (r&1) == 0 && (r&2) != 0 {
		return int(val)
	}
	if (r&1) != 0 && (r&2) != 0 {
		return int(val / 100)
	}
	return 0
}

func (r RKNumber) Float64() float64 {
	val := int32(r) >> 2
	v2 := math.Float64frombits(uint64(val) << 34)

	if (r&1) == 0 && (r&2) == 0 {
		return v2
	}
	if (r&1) != 0 && (r&2) == 0 {
		return v2 / 100.0
	}
	return 0.0
}

func (r RKNumber) String() string {
	if r.IsInteger() {
		return fmt.Sprint(r.Int())
	}
	return fmt.Sprint(r.Float64())
}
