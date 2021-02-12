package simple

import (
	"encoding/csv"
	"os"

	"github.com/pbnjay/grate"
)

var _ = grate.Register("csv", 15, OpenCSV)

// OpenCSV defines a Source's instantiation function.
// It should return ErrNotInFormat immediately if filename is not of the correct file type.
func OpenCSV(filename string) (grate.Source, error) {
	f, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	t := &simpleFile{
		filename: filename,
		iterRow:  -1,
	}

	s := csv.NewReader(f)
	s.FieldsPerRecord = -1

	total := 0
	ncols := make(map[int]int)
	rec, err := s.Read()
	for ; err == nil; rec, err = s.Read() {
		ncols[len(rec)]++
		total++
		t.rows = append(t.rows, rec)
	}
	if err != nil {
		switch perr := err.(type) {
		case *csv.ParseError:
			return nil, grate.WrapErr(perr, grate.ErrNotInFormat)
		}
		if total < 10 {
			// probably? not in this format
			return nil, grate.WrapErr(err, grate.ErrNotInFormat)
		}
		return nil, err
	}

	// kinda arbitrary metrics for detecting CSV
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
