package commonxl

import (
	"errors"
	"fmt"
	"regexp"
	"strings"
)

// Formatter contains formatting methods common to Excel spreadsheets.
type Formatter struct {
	flags           uint64
	customCodes     map[uint16]FmtFunc
	customCodeTypes map[uint16]CellType
}

const (
	fMode1904 uint64 = 1
)

// Mode1904 indicates that dates start on Jan 1, 1904
// this setting was used in early MacOS Excel applications.
func (x *Formatter) Mode1904(enabled bool) {
	if enabled {
		x.flags |= fMode1904
	} else {
		x.flags = x.flags &^ fMode1904
	}
}

// Add a custom number format to the formatter.
func (x *Formatter) Add(fmtID uint16, formatCode string) error {
	if x.customCodes == nil {
		x.customCodes = make(map[uint16]FmtFunc)
		x.customCodeTypes = make(map[uint16]CellType)
	}
	if strings.ToLower(formatCode) == "general" {
		x.customCodes[fmtID] = goFormatters[0]
		return nil
	}
	_, ok := goFormatters[fmtID]
	if ok {
		return errors.New("grate/commonxl: cannot replace default number formats")
	}

	_, ok2 := x.customCodes[fmtID]
	if ok2 {
		return errors.New("grate/commonxl: cannot replace existing number formats")
	}

	x.customCodes[fmtID], x.customCodeTypes[fmtID] = makeFormatter(formatCode)
	return nil
}

func (x *Formatter) getCellType(fmtID uint16) (CellType, bool) {
	if ct, ok := builtInFormatTypes[fmtID]; ok {
		return ct, true
	}
	if x.customCodeTypes != nil {
		ct, ok := x.customCodeTypes[fmtID]
		return ct, ok
	}
	return 0, false
}

var (
	minsMatch = regexp.MustCompile("h.*m.*s*")
	nonEsc    = regexp.MustCompile(`([^"]|^)"`)
	squash    = regexp.MustCompile(`[*_].`)
	fixEsc    = regexp.MustCompile(`\\(.)`)

	formatMatchBrackets    = regexp.MustCompile(`\[[^\]]*\]`)
	formatMatchTextLiteral = regexp.MustCompile(`"[^"]*"`)
)

func makeFormatter(s string) (FmtFunc, CellType) {
	//log.Printf("makeFormatter('%s')", s)
	// remove any coloring marks
	s = formatMatchBrackets.ReplaceAllString(s, "")
	if strings.Contains(s, ";") {
		parts := strings.Split(s, ";")
		posFF, ctypePos := makeFormatter(parts[0])
		rem := make([]FmtFunc, len(parts)-1)
		for i, ps := range parts[1:] {
			rem[i], _ = makeFormatter(ps)
		}
		return switchFmtFunc(posFF, rem...), ctypePos
	}

	// escaped characters, and quoted text
	s2 := fixEsc.ReplaceAllString(s, "")
	s2 = formatMatchTextLiteral.ReplaceAllString(s, "")

	if strings.ContainsAny(s2, "ymdhs") {
		// it's a date/time format

		if loc := minsMatch.FindStringIndex(s); loc != nil {
			// m or mm in loc[0]:loc[1] is a minute format
			inner := s[loc[0]:loc[1]]
			inner = strings.Replace(inner, "mm", "04", 1)
			inner = strings.Replace(inner, "m", "4", 1)
			s = s[:loc[0]] + inner + s[loc[1]:]
		}
		dfreps := [][]string{
			{"hh", "15"}, {"h", "15"},
			{"ss", "05"}, {"s", "5"},
			{"mmmmm", "Jan"}, // super ambiguous, replace with 3-letter month
			{"mmmm", "January"}, {"mmm", "Jan"},
			{"mm", "01"}, {"m", "1"},
			{"dddd", "Monday"}, {"ddd", "Mon"},
			{"dd", "02"}, {"d", "2"},
			{"yyyy", "2006"}, {"yy", "06"},
		}
		if strings.Contains(s, "AM") || strings.Contains(s, "PM") {
			dfreps[0][1] = "03"
			dfreps[1][1] = "3"
		}
		for _, dfr := range dfreps {
			s = strings.Replace(s, dfr[0], dfr[1], 1)
		}

		s = nonEsc.ReplaceAllString(s, `$1`)
		s = squash.ReplaceAllString(s, ``)
		s = fixEsc.ReplaceAllString(s, `$1`)

		//log.Printf("   made time formatter '%s'", s)
		return timeFmtFunc(s), DateCell
	}

	var ff FmtFunc
	var ctype CellType
	if strings.ContainsAny(s, ".Ee") {
		verb := "f"
		if strings.ContainsAny(s, "Ee") {
			verb = "E"
		}
		s = regexp.MustCompile("[eE]+[+-]0+").ReplaceAllString(s, "")
		s2 := strings.ReplaceAll(s, ",", "")
		i1 := strings.IndexAny(s2, "0")
		i2 := strings.IndexByte(s2, '.')
		i3 := strings.LastIndexAny(s2, "0.")
		mul := 1
		if strings.Contains(s2, "%") {
			mul = 100
		}
		sf := fmt.Sprintf("%%%d.%d%s", i3-i1, i3-i2, verb)
		//log.Printf("   made float formatter '%s'", sf)
		ff = sprintfFunc(sf, mul)
		ctype = FloatCell
	} else {
		s2 := strings.ReplaceAll(s, ",", "")
		i1 := strings.IndexAny(s2, "0")
		i2 := strings.LastIndexAny(s2, "0.")
		mul := 1
		if strings.Contains(s2, "%") {
			mul = 100
		}
		sf := fmt.Sprintf("%%%dd", i2-i1)
		if (i2 - i1) == 0 {
			sf = "%d"
		}
		//log.Printf("   made int formatter '%s'", sf)
		ff = sprintfFunc(sf, mul)
		ctype = IntegerCell
	}

	if strings.Contains(s, ",") {
		ff = addCommas(ff)
		//log.Printf("   added commas")
	}

	surReg := regexp.MustCompile(`[0#?,.]+`)
	prepost := surReg.Split(s, 2)
	if len(prepost) > 0 && len(prepost[0]) > 0 {
		prepost[0] = nonEsc.ReplaceAllString(prepost[0], `$1`)
		prepost[0] = squash.ReplaceAllString(prepost[0], ``)
		prepost[0] = fixEsc.ReplaceAllString(prepost[0], `$1`)
	}
	if len(prepost) == 1 {
		if prepost[0] == "@" {
			return identFunc, StringCell
		}
		//log.Printf("   added static ('%s')", prepost[0])
		return staticFmtFunc(prepost[0]), StringCell
	}
	if len(prepost[0]) > 0 || len(prepost[1]) > 0 {
		prepost[1] = nonEsc.ReplaceAllString(prepost[1], `$1`)
		prepost[1] = squash.ReplaceAllString(prepost[1], ``)
		prepost[1] = fixEsc.ReplaceAllString(prepost[1], `$1`)

		ff = surround(prepost[0], ff, prepost[1])
		//log.Printf("   added surround ('%s' ... '%s')", prepost[0], prepost[1])
	}

	return ff, ctype
}

// Get the number format func to use for formatting values,
// it returns false when fmtID is unknown.
func (x *Formatter) Get(fmtID uint16) (FmtFunc, bool) {
	ff, ok := goFormatters[fmtID]
	if !ok {
		fs, ok2 := x.customCodes[fmtID]
		if ok2 {
			return fs, true
		}
		ff = identFunc
	}

	return ff, ok
}

// Apply the specified number format to the value.
// Returns false when fmtID is unknown.
func (x *Formatter) Apply(fmtID uint16, val interface{}) (string, bool) {
	ff, ok := goFormatters[fmtID]
	if !ok {
		fs, ok2 := x.customCodes[fmtID]
		if ok2 {
			return fs(x, val), true
		}
	}
	return ff(x, val), ok
}

// builtInFormats are all the built-in number formats for XLS/XLSX.
var builtInFormats = map[uint16]string{
	0:  `General`,
	1:  `0`,
	2:  `0.00`,
	3:  `#,##0`,
	4:  `#,##0.00`,
	9:  `0%`,
	10: `0.00%`,

	11: `0.00E+00`,
	12: `# ?/?`,
	13: `# ??/??`,
	14: `mm-dd-yy`,
	15: `d-mmm-yy`,
	16: `d-mmm`,
	17: `mmm-yy`,
	18: `h:mm AM/PM`,
	19: `h:mm:ss AM/PM`,
	20: `h:mm`,
	21: `h:mm:ss`,
	22: `m/d/yy h:mm`,
	37: `#,##0 ;(#,##0)`,
	38: `#,##0 ;[Red](#,##0)`,
	39: `#,##0.00;(#,##0.00)`,
	40: `#,##0.00;[Red](#,##0.00)`,

	41: `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`,
	42: `_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)`,
	43: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`,
	44: `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`,

	45: `mm:ss`,
	46: `[h]:mm:ss`,
	47: `mmss.0`,
	48: `##0.0E+0`,
	49: `@`,

	// zh-cn format codes
	27: `yyyy"年"m"月"`,
	28: `m"月"d"日"`,
	29: `m"月"d"日"`,
	30: `m-d-yy`,
	31: `yyyy"年"m"月"d"日"`,
	32: `h"时"mm"分"`,
	33: `h"时"mm"分"ss"秒"`,
	34: `上午/下午 h"时"mm"分"`,
	35: `上午/下午 h"时"mm"分"ss"秒"`,
	36: `yyyy"年"m"月"`,
	50: `yyyy"年"m"月"`,
	51: `m"月"d"日"`,
	52: `yyyy"年"m"月"`,
	53: `m"月"d"日"`,
	54: `m"月"d"日"`,
	55: `上午/下午 h"时"mm"分"`,
	56: `上午/下午 h"时"mm"分"ss"秒`,
	57: `yyyy"年"m"月"`,
	58: `m"月"d"日"`,

	// th-th format codes (in the spec these have a "t" prefix?)
	59: `0`,
	60: `0.00`,
	61: `#,##0`,
	62: `#,##0.00`,
	67: `0%`,
	68: `0.00%`,
	69: `# ?/?`,
	70: `# ??/??`,

	// th format code, but translated to aid the parser
	71: `d/m/yyyy`,      // `ว/ด/ปปปป`,
	72: `d-mmm-yy`,      // `ว-ดดด-ปป`,
	73: `d-mmm`,         // `ว-ดดด`,
	74: `mmm-yy`,        // `ดดด-ปป`,
	75: `h:mm`,          // `ช:นน`,
	76: `h:mm:ss`,       // `ช:นน:ทท`,
	77: `d/m/yyyy h:mm`, // `ว/ด/ปปปป ช:นน`,
	78: `mm:ss`,         // `นน:ทท`,
	79: `[h]:mm:ss`,     // `[ช]:นน:ทท`,
	80: `mm:ss.0`,       // `นน:ทท.0`,
	81: `d/m/bb`,        // `d/m/bb`,
}

// builtInFormatTypes are the underlying datatypes for built-in number formats in XLS/XLSX.
var builtInFormatTypes = map[uint16]CellType{
	// 0 has no defined type
	1:  IntegerCell,
	2:  FloatCell,
	3:  IntegerCell,
	4:  FloatCell,
	9:  FloatCell,
	10: FloatCell,

	11: FloatCell,
	12: FloatCell,
	13: FloatCell,
	14: DateCell,
	15: DateCell,
	16: DateCell,
	17: DateCell,
	18: DateCell,
	19: DateCell,
	20: DateCell,
	21: DateCell,
	22: DateCell,
	37: IntegerCell,
	38: IntegerCell,
	39: FloatCell,
	40: FloatCell,
	41: IntegerCell,
	42: IntegerCell,
	43: FloatCell,
	44: FloatCell,
	45: DateCell, // Durations?
	46: DateCell,
	47: DateCell,
	48: FloatCell,
	49: StringCell,
	27: DateCell,
	28: DateCell,
	29: DateCell,
	30: DateCell,
	31: DateCell,
	32: DateCell,
	33: DateCell,
	34: DateCell,
	35: DateCell,
	36: DateCell,
	50: DateCell,
	51: DateCell,
	52: DateCell,
	53: DateCell,
	54: DateCell,
	55: DateCell,
	56: DateCell,
	57: DateCell,
	58: DateCell,
	59: IntegerCell,
	60: FloatCell,
	61: IntegerCell,
	62: FloatCell,
	67: FloatCell,
	68: FloatCell,
	69: FloatCell,
	70: FloatCell,
	71: DateCell,
	72: DateCell,
	73: DateCell,
	74: DateCell,
	75: DateCell,
	76: DateCell,
	77: DateCell,
	78: DateCell,
	79: DateCell,
	80: DateCell,
	81: DateCell,
}
