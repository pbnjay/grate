package xls

import (
	"os"
	"strings"
	"testing"
)

var testFiles = []string{
	"../testdata/test.xls",
	"../testdata/test2.xls",
	"../testdata/test3.xls",
	"../testdata/test4.xls",
	"../testdata/basic.xls",
	"../testdata/basic2.xls",
}

func TestLoading(t *testing.T) {
	for _, fn := range testFiles {
		wb, err := Open(fn)
		if err != nil {
			t.Fatal(err)
		}

		sheets, err := wb.List()
		if err != nil {
			t.Fatal(err)
		}
		for _, s := range sheets {
			sheet, err := wb.Get(s)
			if err != nil {
				t.Fatal(err)
			}

			for sheet.Next() {
				sheet.Strings()
			}
		}

		err = wb.Close()
		if err != nil {
			t.Fatal(err)
		}
	}
}

func noTestBasic(t *testing.T) {
	trueFile, err := os.ReadFile("../testdata/basic.tsv")
	if err != nil {
		t.Skip()
	}
	lines := strings.Split(string(trueFile), "\n")

	for _, fn := range testFiles {
		wb, err := Open(fn)
		if err != nil {
			t.Fatal(err)
		}

		sheets, err := wb.List()
		if err != nil {
			t.Fatal(err)
		}
		for _, s := range sheets {
			sheet, err := wb.Get(s)
			if err != nil {
				t.Fatal(err)
			}

			i := 0
			for sheet.Next() {
				row := strings.Join(sheet.Strings(), "\t")
				if lines[i] != row {
					t.Fatalf("line %d mismatch: '%s' <> '%s'", i, row, lines[i])
				}
				i++
			}
		}

		err = wb.Close()
		if err != nil {
			t.Fatal(err)
		}
	}
}
