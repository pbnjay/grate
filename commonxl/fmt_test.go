package commonxl

import (
	"log"
	"testing"
	"time"
)

type testcaseNums struct {
	v interface{}
	s string
}

var commas = []testcaseNums{
	{10, "10"},
	{float64(10), "10"},
	{float64(10) + 0.12345, "10.12345"},
	{-10, "-10"},
	{float64(-10), "-10"},
	{float64(-10) + 0.12345, "-9.87655"},
	{uint16(10), "10"},
	{100, "100"},
	{float64(100), "100"},
	{float64(100) + 0.12345, "100.12345"},
	{-100, "-100"},
	{float64(-100), "-100"},
	{float64(-100) + 0.12345, "-99.87655"},
	{uint16(100), "100"},
	{1000, "1,000"},
	{float64(1000), "1,000"},
	{float64(1000) + 0.12345, "1,000.12345"},
	{-1000, "-1,000"},
	{float64(-1000), "-1,000"},
	{float64(-1000) + 0.12345, "-999.87655"},
	{uint16(1000), "1,000"},
	{10000, "10,000"},
	{float64(10000), "10,000"},
	{float64(10000) + 0.12345, "10,000.12345"},
	{-10000, "-10,000"},
	{float64(-10000), "-10,000"},
	{float64(-10000) + 0.12345, "-9,999.87655"},
	{uint16(10000), "10,000"},
	{100000, "100,000"},
	{float64(100000), "100,000"},
	{float64(100000) + 0.12345, "100,000.12345"},
	{-100000, "-100,000"},
	{float64(-100000), "-100,000"},
	{float64(-100000) + 0.12345, "-99,999.87655"},
	{uint64(100000), "100,000"},
	{1000000, "1,000,000"},
	{float64(1000000), "1e+06"},
	{float64(1000000) + 0.12345, "1.00000012345e+06"},
	{-1000000, "-1,000,000"},
	{float64(-1000000), "-1e+06"},
	{float64(-1000000) + 0.12345, "-999,999.87655"},
	{uint64(1000000), "1,000,000"},
	{10000000, "10,000,000"},
	{float64(10000000), "1e+07"},
	{float64(10000000) + 0.12345, "1.000000012345e+07"},
	{-10000000, "-10,000,000"},
	{float64(-10000000), "-1e+07"},
	{float64(-10000000) + 0.12345, "-9.99999987655e+06"},
	{uint64(10000000), "10,000,000"},
	{100000000, "100,000,000"},
	{float64(100000000), "1e+08"},
	{float64(100000000) + 0.12345, "1.0000000012345e+08"},
	{-100000000, "-100,000,000"},
	{float64(-100000000), "-1e+08"},
	{float64(-100000000) + 0.12345, "-9.999999987655e+07"},
	{uint64(100000000), "100,000,000"},
}

func TestCommas(t *testing.T) {
	cf := addCommas(identFunc)
	for _, c := range commas {
		fs := cf(nil, c.v)
		if c.s != fs {
			t.Fatalf("commas failed: get '%s' but expected '%s' for %T(%v)",
				fs, c.s, c.v, c.v)
		}
	}
}

func TestDateFormats(t *testing.T) {
	var testDates = []time.Time{
		time.Date(1901, 7, 11, 1, 5, 0, 0, time.UTC),
		time.Date(1905, 7, 11, 4, 10, 0, 0, time.UTC),
		time.Date(1904, 7, 11, 8, 15, 0, 0, time.UTC),
		time.Date(1993, 7, 11, 12, 20, 0, 0, time.UTC),
		time.Date(1983, 7, 11, 16, 30, 0, 0, time.UTC),
		time.Date(1983, 7, 11, 20, 45, 0, 0, time.UTC),
		time.Date(2000, 12, 31, 23, 59, 0, 0, time.UTC),
		time.Date(2002, 12, 31, 23, 59, 0, 0, time.UTC),
		time.Date(2012, 3, 10, 9, 30, 0, 0, time.UTC),
		time.Date(2014, 3, 27, 9, 37, 0, 0, time.UTC),
	}

	fx := &Formatter{}
	for _, t := range testDates {
		for fid, ctype := range builtInFormatTypes {
			if ctype != DateCell {
				continue
			}
			ff, _ := goFormatters[fid]
			// mainly testing these don't crash...
			log.Println(ff(fx, t))
		}
	}
}
func TestBoolFormats(t *testing.T) {
	ff, _ := makeFormatter(`"yes";"yes";"no"`)

	if "no" != ff(nil, false) {
		t.Fatal(`false should be "no"`)
	}
	if "no" != ff(nil, 0) {
		t.Fatal(`0 should be "no"`)
	}
	if "no" != ff(nil, 0.0) {
		t.Fatal(`0.0 should be "no"`)
	}

	/////

	if "yes" != ff(nil, true) {
		t.Fatal(`true should be "yes"`)
	}
	if "yes" != ff(nil, 99) {
		t.Fatal(`99 should be "yes"`)
	}
	if "yes" != ff(nil, -4) {
		t.Fatal(`-4 should be "yes"`)
	}

	if "yes" != ff(nil, 4.0) {
		t.Fatal(`4.0 should be "yes"`)
	}
	if "yes" != ff(nil, -99.0) {
		t.Fatal(`-99.0 should be "yes"`)
	}
}
