package commonxl

import (
	"log"
	"strings"
	"time"
)

// ConvertToDate converts a floating-point value using the
// Excel date serialization conventions.
func (x *Formatter) ConvertToDate(val float64) time.Time {
	// http://web.archive.org/web/20190808062235/http://aa.usno.navy.mil/faq/docs/JD_Formula.php
	v := int(val)
	if v < 61 {
		jdate := val + 0.5
		if (x.flags & fMode1904) != 0 {
			jdate += 2416480.5
		} else {
			jdate += 2415018.5
		}
		JD := int(jdate)
		frac := jdate - float64(JD)

		L := JD + 68569
		N := 4 * L / 146097
		L = L - (146097*N+3)/4
		I := 4000 * (L + 1) / 1461001
		L = L - 1461*I/4 + 31
		J := 80 * L / 2447
		day := L - 2447*J/80
		L = J / 11
		month := time.Month(J + 2 - 12*L)
		year := 100*(N-49) + I + L

		t := time.Duration(float64(time.Hour*24) * frac)
		return time.Date(year, month, day, 0, 0, 0, 0, time.UTC).Add(t)
	}
	frac := val - float64(v)
	date := time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
	if (x.flags & fMode1904) == 0 {
		date = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	}

	t := time.Duration(float64(time.Hour*24) * frac)
	return date.AddDate(0, 0, v).Add(t)
}

func timeFmtFunc(f string) FmtFunc {
	return func(x *Formatter, v interface{}) string {
		t, ok := v.(time.Time)
		if !ok {
			fval, ok := convertToFloat64(v)
			if !ok {
				return "MUST BE time.Time OR numeric TO FORMAT CORRECTLY"
			}
			t = x.ConvertToDate(fval)
		}
		log.Println("formatting date", t, "with", f, "=", t.Format(f))
		return t.Format(f)
	}
}

func cnTimeFmtFunc(f string) FmtFunc {
	return func(x *Formatter, v interface{}) string {
		t, ok := v.(time.Time)
		if !ok {
			fval, ok := convertToFloat64(v)
			if !ok {
				return "MUST BE time.Time OR numeric TO FORMAT CORRECTLY"
			}
			t = x.ConvertToDate(fval)
		}
		s := t.Format(f)
		s = strings.Replace(s, `AM`, `上午`, 1)
		return strings.Replace(s, `PM`, `下午`, 1)
	}
}

// 0x0001 = date   0b0010 = time    0b0011 = date+time
var builtInDateFormats = map[uint16]byte{
	14: 1, 15: 1, 16: 1, 17: 1, 18: 2, 19: 2, 20: 2, 21: 2, 22: 3,
	45: 2, 46: 2, 47: 2, 27: 1, 28: 1, 29: 1, 30: 1, 31: 1, 32: 2,
	33: 2, 34: 2, 35: 2, 36: 1, 50: 1, 51: 1, 52: 1, 53: 1, 54: 1,
	55: 2, 56: 2, 57: 1, 58: 1, 71: 1, 72: 1, 73: 1, 74: 1, 75: 2,
	76: 2, 77: 3, 78: 2, 79: 2, 80: 2, 81: 1,
}
