package commonxl

import (
	"fmt"
	"strconv"
	"strings"
)

// FmtFunc will format a value according to the designated style.
type FmtFunc func(*Formatter, interface{}) string

func staticFmtFunc(s string) FmtFunc {
	return func(x *Formatter, v interface{}) string {
		return s
	}
}

func surround(pre string, ff FmtFunc, post string) FmtFunc {
	return func(x *Formatter, v interface{}) string {
		return pre + ff(x, v) + post
	}
}

func addNegParens(ff FmtFunc) FmtFunc {
	return func(x *Formatter, v interface{}) string {
		s1 := ff(x, v)
		if s1[0] == '-' {
			return "(" + s1[1:] + ")"
		}
		return s1
	}
}

func addCommas(ff FmtFunc) FmtFunc {
	return func(x *Formatter, v interface{}) string {
		s1 := ff(x, v)
		isNeg := false
		if s1[0] == '-' {
			isNeg = true
			s1 = s1[1:]
		}
		endIndex := strings.IndexAny(s1, ".eE")
		if endIndex < 0 {
			endIndex = len(s1)
		}
		for endIndex > 3 {
			endIndex -= 3
			s1 = s1[:endIndex] + "," + s1[endIndex:]
		}
		if isNeg {
			return "-" + s1
		}
		return s1
	}
}

func identFunc(x *Formatter, v interface{}) string {
	if s, ok := v.(string); ok {
		return s
	}
	return fmt.Sprint(v)
}

func sprintfFunc(fs string, mul int) FmtFunc {
	wantInt64 := strings.Contains(fs, "%d")
	return func(x *Formatter, v interface{}) string {
		switch val := v.(type) {
		case int, uint, int64, uint64, int32, uint32, uint16, int16:
			return fmt.Sprintf(fs, v)

		case float64:
			val *= float64(mul)
			if wantInt64 {
				v2 := int64(val)
				return fmt.Sprintf(fs, v2)
			}
			return fmt.Sprintf(fs, val)
		}
		return fmt.Sprint(v)
	}
}

func convertToInt64(v interface{}) (int64, bool) {
	x, ok := convertToFloat64(v)
	return int64(x), ok
}

func convertToFloat64(v interface{}) (float64, bool) {
	switch val := v.(type) {
	case float64:
		return val, true
	case bool:
		if val {
			return 1.0, true
		}
		return 0.0, true
	case int:
		return float64(val), true
	case int8:
		return float64(val), true
	case int16:
		return float64(val), true
	case int32:
		return float64(val), true
	case int64:
		return float64(val), true
	case uint:
		return float64(val), true
	case uint8:
		return float64(val), true
	case uint16:
		return float64(val), true
	case uint32:
		return float64(val), true
	case uint64:
		return float64(val), true
	case float32:
		return float64(val), true
	case string:
		nf, err := strconv.ParseFloat(val, 64)
		return nf, err == nil
	default:
		return 0.0, false
	}
}

func zeroDashFunc(ff FmtFunc) FmtFunc {
	return func(x *Formatter, v interface{}) string {
		fval, ok := convertToFloat64(v)
		if !ok {
			// strings etc returned as-is
			return fmt.Sprint(v)
		}
		if fval == 0.0 {
			return "-"
		}
		return ff(x, v)
	}
}

func fracFmtFunc(n int) FmtFunc {
	return func(x *Formatter, v interface{}) string {
		f, ok := convertToFloat64(v)
		if !ok {
			return "MUST BE numeric TO FORMAT CORRECTLY"
		}
		w, n, d := DecimalToWholeFraction(f, n, n)
		if n == 0 {
			return fmt.Sprintf("%d", w)
		}
		return fmt.Sprintf("%d %d/%d", w, n, d)
	}
}

// handle (up to) all four format cases
func switchFmtFunc(pos FmtFunc, others ...FmtFunc) FmtFunc {
	stringFF := identFunc
	zeroFF := pos
	negFF := pos
	if len(others) > 0 {
		negFF = others[0]
		if len(others) > 1 {
			zeroFF = others[1]
			if len(others) > 2 {
				stringFF = others[2]
			}
		}
	}
	return func(x *Formatter, v interface{}) string {
		val, ok := convertToFloat64(v)
		if !ok {
			return stringFF(x, v)
		}
		if val == 0.0 {
			return zeroFF(x, v)
		}
		if val < 0.0 {
			return negFF(x, v)
		}
		return pos(x, v)
	}
}

// mapping of standard built-ins to Go date format funcs.
var goFormatters = map[uint16]FmtFunc{
	0:  identFunc, // FIXME: better "general" formatter
	49: identFunc,

	14: timeFmtFunc(`01-02-06`),
	15: timeFmtFunc(`2-Jan-06`),
	16: timeFmtFunc(`2-Jan`),
	17: timeFmtFunc(`Jan-06`),
	20: timeFmtFunc(`15:04`),
	21: timeFmtFunc(`15:04:05`),
	22: timeFmtFunc(`1/2/06 15:04`),
	45: timeFmtFunc(`04:05`),
	46: timeFmtFunc(`3:04:05`),
	47: timeFmtFunc(`0405.9`),
	27: timeFmtFunc(`2006"年"1"月"`),
	28: timeFmtFunc(`1"月"2"日"`),
	29: timeFmtFunc(`1"月"2"日"`),
	30: timeFmtFunc(`1-2-06`),
	31: timeFmtFunc(`2006"年"1"月"2"日"`),
	32: timeFmtFunc(`15"时"04"分"`),
	33: timeFmtFunc(`15"时"04"分"05"秒"`),
	36: timeFmtFunc(`2006"年"2"月"`),
	50: timeFmtFunc(`2006"年"2"月"`),
	51: timeFmtFunc(`1"月"2"日"`),
	52: timeFmtFunc(`2006"年"1"月"`),
	53: timeFmtFunc(`1"月"2"日"`),
	54: timeFmtFunc(`1"月"2"日"`),
	57: timeFmtFunc(`2006"年"1"月"`),
	58: timeFmtFunc(`1"月"2"日"`),
	71: timeFmtFunc(`2/1/2006`),
	72: timeFmtFunc(`2-Jan-06`),
	73: timeFmtFunc(`2-Jan`),
	74: timeFmtFunc(`Jan-06`),
	75: timeFmtFunc(`15:04`),
	76: timeFmtFunc(`15:04:05`),
	77: timeFmtFunc(`2/1/2006 15:04`),
	78: timeFmtFunc(`04:05`),
	79: timeFmtFunc(`15:04:05`),
	80: timeFmtFunc(`04:05.9`),
	81: timeFmtFunc(`2/1/06`),
	18: timeFmtFunc(`3:04 PM`),
	19: timeFmtFunc(`3:04:05 PM`),

	34: cnTimeFmtFunc(`PM 3"时"04"分"`),
	35: cnTimeFmtFunc(`PM 3"时"04"分"05"秒"`),
	55: cnTimeFmtFunc(`PM 3"时"04"分"`),
	56: cnTimeFmtFunc(`PM 3"时"04"分"05"秒`),

	12: fracFmtFunc(1),
	13: fracFmtFunc(2),

	69: surround("t", fracFmtFunc(1), ""),
	70: surround("t", fracFmtFunc(2), ""),

	1:  sprintfFunc(`%d`, 1),
	2:  sprintfFunc(`%4.2f`, 1),
	59: sprintfFunc(`t%d`, 1),
	60: sprintfFunc(`t%4.2f`, 1),

	9:  sprintfFunc(`%d%%`, 100),
	10: sprintfFunc(`%4.2f%%`, 100),
	67: sprintfFunc(`t%d%%`, 100),
	68: sprintfFunc(`t%4.2f%%`, 100),

	3:  addCommas(sprintfFunc("%d", 1)),
	61: surround("t", addCommas(sprintfFunc("%d", 1)), ""),
	37: addNegParens(addCommas(sprintfFunc("%d", 1))),
	38: addNegParens(addCommas(sprintfFunc("%d", 1))),

	4:  addCommas(sprintfFunc("%4.2f", 1)),
	62: surround("t", addCommas(sprintfFunc("%4.2f", 1)), ""),
	39: addNegParens(addCommas(sprintfFunc("%4.2f", 1))),
	40: addNegParens(addCommas(sprintfFunc("%4.2f", 1))),

	11: sprintfFunc(`%4.2E`, 1),
	48: sprintfFunc(`%3.1E`, 1),

	41: zeroDashFunc(addCommas(sprintfFunc("%d", 1))),
	43: zeroDashFunc(addCommas(sprintfFunc("%4.2f", 1))),

	42: switchFmtFunc(
		surround("$", addCommas(sprintfFunc("%d", 1)), ""),
		surround("$(", addCommas(sprintfFunc("%d", 1)), ")"),
		staticFmtFunc("$-")),
	44: switchFmtFunc(
		surround("$", addCommas(sprintfFunc("%4.2f", 1)), ""),
		surround("$(", addCommas(sprintfFunc("%4.2f", 1)), ")"),
		staticFmtFunc("$-")),
}
