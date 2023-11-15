package xls

import (
	"testing"

	"github.com/pbnjay/grate"
)

func TestDecimalNumberSavedAsIntegerMultipliedByHundred(t *testing.T) {
	wb, _ := grate.Open("../testdata/decimal.xls")
	sheets, _ := wb.List()
	for _, s := range sheets {
		sheet, _ := wb.Get(s)
		sheet.Next()

		var value float64
		sheet.Scan(&value)

		if value != 1.75 {
			t.Log("Expected value to be 1.75, but actually is", value)
			t.Fail()
		}
	}
	wb.Close()
}
