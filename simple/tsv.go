package simple

import (
	"bufio"
	"os"
	"strings"

	"github.com/pbnjay/grate"
)

var _ = grate.Register("tsv", 10, OpenTSV)

// OpenTSV defines a Source's instantiation function.
// It should return ErrNotInFormat immediately if filename is not of the correct file type.
func OpenTSV(filename string) (grate.Source, error) {
	f, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	t := &simpleFile{
		filename: filename,
		iterRow:  -1,
	}

	s := bufio.NewScanner(f)
	total := 0
	ncols := make(map[int]int)
	for s.Scan() {
		r := strings.Split(s.Text(), "\t")
		ncols[len(r)]++
		total++
		t.rows = append(t.rows, r)
	}

	// kinda arbitrary metrics for detecting TSV
	looksGood := 0
	for c, n := range ncols {
		if c <= 1 {
			continue
		}
		if n > 10 && float64(n)/float64(total) > 0.8 {
			// more than 80% of rows have the same number of columns, we're good
			looksGood = 2
		} else if n > 25 && looksGood == 0 {
			looksGood = 1
		}
	}
	if looksGood == 1 {
		return t, grate.ErrNotInFormat
	}

	return t, nil
}
