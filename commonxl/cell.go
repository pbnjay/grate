package commonxl

import (
	"fmt"
	"math"
	"net/url"
	"strconv"
	"time"
	"unicode/utf16"
)

// CellType annotates the type of data extracted in the cell.
type CellType uint16

// CellType annotations for various cell value types.
const (
	BlankCell CellType = iota
	IntegerCell
	FloatCell
	StringCell
	BooleanCell
	DateCell

	HyperlinkStringCell // internal type to separate URLs
	StaticCell          // placeholder, internal use only
)

// String returns a string description of the cell data type.
func (c CellType) String() string {
	switch c {
	case BlankCell:
		return "blank"
	case IntegerCell:
		return "integer"
	case FloatCell:
		return "float"
	case BooleanCell:
		return "boolean"
	case DateCell:
		return "date"
	case HyperlinkStringCell:
		return "hyperlink"
	default: // StringCell, StaticCell
		return "string"
	}
}

// Cell represents a single cell value.
type Cell []interface{}

// internally, it is a slice sized 2 or 3
//   [Value, CellType] or [Value, CellType, FormatNumber]
// where FormatNumber is a uint16 if not 0

// Value returns the contents as a generic interface{}.
func (c Cell) Value() interface{} {
	if len(c) == 0 {
		return ""
	}
	return c[0]
}

// SetURL adds a URL hyperlink to the cell.
func (c *Cell) SetURL(link string) {
	(*c)[1] = HyperlinkStringCell
	if len(*c) == 2 {
		*c = append(*c, uint16(0), link)
	} else { // len = 3 already
		*c = append(*c, link)
	}
}

// URL returns the parsed URL when a cell contains a hyperlink.
func (c Cell) URL() (*url.URL, bool) {
	if c.Type() == HyperlinkStringCell && len(c) >= 4 {
		u, err := url.Parse(c[3].(string))
		return u, err == nil
	}
	return nil, false
}

// Type returns the CellType of the value.
func (c Cell) Type() CellType {
	if len(c) < 2 {
		return BlankCell
	}
	return c[1].(CellType)
}

// FormatNo returns the NumberFormat used for display.
func (c Cell) FormatNo() uint16 {
	if len(c) == 3 {
		return c[2].(uint16)
	}
	return 0
}

///////

var boolStrings = map[string]bool{
	"yes": true, "true": true, "t": true, "y": true, "1": true, "on": true,
	"no": false, "false": false, "f": false, "n": false, "0": false, "off": false,
	"YES": true, "TRUE": true, "T": true, "Y": true, "1.0": true, "ON": true,
	"NO": false, "FALSE": false, "F": false, "N": false, "0.0": false, "OFF": false,
}

// NewCellWithType creates a new cell value with the given type, coercing as necessary.
func NewCellWithType(value interface{}, t CellType, f *Formatter) Cell {
	c := NewCell(value)
	if c[1] == t {
		// fast path if it was already typed correctly
		return c
	}

	if c[1] == BooleanCell {
		if t == IntegerCell {
			if c[0].(bool) {
				c[0] = int64(1)
			} else {
				c[0] = int64(0)
			}
			c[1] = IntegerCell
		} else if t == FloatCell {
			if c[0].(bool) {
				c[0] = float64(1.0)
			} else {
				c[0] = float64(0.0)
			}
			c[1] = FloatCell
		} else if t == StringCell {
			if c[0].(bool) {
				c[0] = "TRUE"
			} else {
				c[0] = "FALSE"
			}
			c[1] = FloatCell
		}
	}

	if c[1] == FloatCell {
		if t == IntegerCell {
			c[0] = int64(c[0].(float64))
			c[1] = IntegerCell
		} else if t == BooleanCell {
			c[0] = c[0].(float64) != 0.0
			c[1] = BooleanCell
		}
	}
	if c[1] == IntegerCell {
		if t == FloatCell {
			c[0] = float64(c[0].(int64))
			c[1] = FloatCell
		} else if t == BooleanCell {
			c[0] = c[0].(int64) != 0
			c[1] = BooleanCell
		}
	}
	if c[1] == StringCell {
		if t == IntegerCell {
			x, _ := strconv.ParseInt(c[0].(string), 10, 64)
			c[0] = x
			c[1] = IntegerCell
		} else if t == FloatCell {
			x, _ := strconv.ParseFloat(c[0].(string), 64)
			c[0] = x
			c[1] = FloatCell
		} else if t == BooleanCell {
			c[0] = boolStrings[c[0].(string)]
			c[1] = BooleanCell
		}
	}
	if t == StringCell {
		c[0] = fmt.Sprint(c[0])
		c[1] = StringCell
	}
	if t == DateCell {
		if c[1] == FloatCell {
			c[0] = f.ConvertToDate(c[0].(float64))
		} else if c[1] == IntegerCell {
			c[0] = f.ConvertToDate(float64(c[0].(int64)))
		}
		c[1] = DateCell
	}
	return c
}

// NewCell creates a new cell value from any builtin type.
func NewCell(value interface{}) Cell {
	c := make([]interface{}, 2)
	switch v := value.(type) {
	case bool:
		c[0] = v
		c[1] = BooleanCell
	case int:
		c[0] = int64(v)
		c[1] = IntegerCell
	case int8:
		c[0] = int64(v)
		c[1] = IntegerCell
	case int16:
		c[0] = int64(v)
		c[1] = IntegerCell
	case int32:
		c[0] = int64(v)
		c[1] = IntegerCell
	case int64:
		c[0] = int64(v)
		c[1] = IntegerCell
	case uint8:
		c[0] = int64(v)
		c[1] = IntegerCell
	case uint16:
		c[0] = int64(v)
		c[1] = IntegerCell
	case uint32:
		c[0] = int64(v)
		c[1] = IntegerCell

	case uint:
		if v > math.MaxInt64 {
			c[0] = float64(v)
			c[1] = FloatCell
		} else {
			c[0] = int64(v)
			c[1] = IntegerCell
		}
	case uint64:
		if v > math.MaxInt64 {
			c[0] = float64(v)
			c[1] = FloatCell
		} else {
			c[0] = int64(v)
			c[1] = IntegerCell
		}

	case float32:
		c[0] = float64(v)
		c[1] = FloatCell
	case float64:
		c[0] = float64(v)
		c[1] = FloatCell

	case string:
		if len(v) == 0 {
			c[0] = nil
			c[1] = BlankCell
		} else {
			c[0] = v
			c[1] = StringCell
		}
	case []byte:
		if len(v) == 0 {
			c[0] = nil
			c[1] = BlankCell
		} else {
			c[0] = string(v)
			c[1] = StringCell
		}
	case []uint16:
		if len(v) == 0 {
			c[0] = nil
			c[1] = BlankCell
		} else {
			c[0] = string(utf16.Decode(v))
			c[1] = StringCell
		}
	case []rune:
		if len(v) == 0 {
			c[0] = nil
			c[1] = BlankCell
		} else {
			c[0] = string(v)
			c[1] = StringCell
		}
	case time.Time:
		c[0] = v
		c[1] = DateCell

	case fmt.Stringer:
		s := v.String()
		if len(s) == 0 {
			c[0] = nil
			c[1] = BlankCell
		} else {
			c[0] = s
			c[1] = StringCell
		}
	default:
		panic("grate: data type not handled")
	}
	return Cell(c)
}

// SetFormatNumber changes the number format stored with the cell.
func (c *Cell) SetFormatNumber(f uint16) {
	if f == 0 {
		*c = (*c)[:2]
		return
	}

	if len(*c) == 2 {
		*c = append(*c, f)
	} else {
		(*c)[2] = f
	}
}
