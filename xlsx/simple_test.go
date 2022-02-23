package xlsx

import (
	"bufio"
	"log"
	"os"
	"strings"
	"testing"

	"github.com/pbnjay/grate/commonxl"
)

var testFilePairs = [][]string{
	{"../testdata/basic.xlsx", "../testdata/basic.tsv"},

	// TODO: custom formatter support
	//{"../testdata/basic2.xlsx", "../testdata/basic2.tsv"},

	// TODO: datetime and fraction formatter support
	//{"../testdata/multi_test.xlsx", "../testdata/multi_test.tsv"},
}

func loadTestData(fn string, ff *commonxl.Formatter) (*commonxl.Sheet, error) {
	f, err := os.Open(fn)
	if err != nil {
		return nil, err
	}
	xs := &commonxl.Sheet{
		Formatter: ff,
	}

	row := 0
	s := bufio.NewScanner(f)
	for s.Scan() {
		record := strings.Split(s.Text(), "\t")
		for i, val := range record {
			xs.Put(row, i, val, 0)
		}
		row++
	}
	return xs, f.Close()
}

func TestBasic(t *testing.T) {
	for _, fnames := range testFilePairs {
		var trueData *commonxl.Sheet
		log.Println("Testing ", fnames[0])

		wb, err := Open(fnames[0])
		if err != nil {
			t.Fatal(err)
		}

		sheets, err := wb.List()
		if err != nil {
			t.Fatal(err)
		}
		firstLoad := true
		for _, s := range sheets {
			sheet, err := wb.Get(s)
			if err != nil {
				t.Fatal(err)
			}
			xsheet := sheet.(*commonxl.Sheet)
			if firstLoad {
				trueData, err = loadTestData(fnames[1], xsheet.Formatter)
				if err != nil {
					t.Fatal(err)
				}
				firstLoad = false
			}

			for xrow, xdata := range xsheet.Rows {
				for xcol, xval := range xdata {
					//t.Logf("at %s (%d,%d) expect '%v'", fnames[0], xrow, xcol, trueData.Rows[xrow][xcol])
					if !trueData.Rows[xrow][xcol].Equal(xval) {
						t.Logf("mismatch at %s (%d,%d): '%v' <> '%v' expected", fnames[0], xrow, xcol,
							xval, trueData.Rows[xrow][xcol])
						t.Fail()
					}
				}
			}
		}

		err = wb.Close()
		if err != nil {
			t.Fatal(err)
		}
	}
}
