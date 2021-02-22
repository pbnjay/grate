package xlsx

import (
	"os"
	"strings"
	"testing"
)

var testFiles = []string{
	"../testdata/multi_test.xlsx",
	"../testdata/basic.xlsx",
	"../testdata/basic2.xlsx",
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

func TestBasic(t *testing.T) {
	trueFile, err := os.ReadFile("../testdata/basic.tsv")
	if err != nil {
		t.Skip()
	}
	lines := strings.Split(string(trueFile), "\n")

	fn := "../testdata/basic.xlsx"
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

func TestBasic2(t *testing.T) {
	trueFile, err := os.ReadFile("../testdata/basic2.tsv")
	if err != nil {
		t.Skip()
	}
	lines := strings.Split(string(trueFile), "\n")

	fn := "../testdata/basic2.xlsx"
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

func TestMulti(t *testing.T) {
	trueFile, err := os.ReadFile("../testdata/multi_test.tsv")
	if err != nil {
		t.Skip()
	}
	lines := strings.Split(string(trueFile), "\n")

	fn := "../testdata/multi_test.xlsx"
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
