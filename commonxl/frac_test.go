package commonxl

import (
	"math"
	"testing"
)

type testcaseFrac struct {
	v float64
	s string
	n int
}

var fracs = []testcaseFrac{
	{10, "10", 1},
	{-10, "-10", 1},
	{10.5, "10 1/2", 1},
	{-10.5, "-10 1/2", 1},

	{10.25, "10 1/4", 1},
	{10.75, "10 3/4", 1},
	{10.667, "10 2/3", 1},

	{-10.25, "-10 1/4", 1},
	{-10.75, "-10 3/4", 1},
	{-10.667, "-10 2/3", 1},

	{3.14159, "3 1/7", 1},
	{3.14159, "3 1/7", 2},
	{3.14159, "3 16/113", 3},
	{3.14159, "3 431/3044", 4},
	{3.14159, "3 3432/24239", 5},
	{3.14159, "3 14159/100000", 6},

	{math.Pi, "3 1/7", 1},
	{math.Pi, "3 1/7", 2},
	{math.Pi, "3 16/113", 3}, // err = 2.6e-7
	{math.Pi, "3 16/113", 4}, // better because 431/3044 err = 2.6e-6
	{math.Pi, "3 14093/99532", 5},
	{math.Pi, "3 14093/99532", 6},

	{-math.Pi, "-3 1/7", 1},
	{-math.Pi, "-3 1/7", 2},
	{-math.Pi, "-3 16/113", 3}, // err = 2.6e-7
	{-math.Pi, "-3 16/113", 4}, // better because 431/3044 err = 2.6e-6
	{-math.Pi, "-3 14093/99532", 5},
	{-math.Pi, "-3 14093/99532", 6},

	// TODO: fixed denominator fractions (e.g. "??/8" )
	// TODO: string interpolations (e.g. '0 "pounds and " ??/100 "pence"')
	// examples: https://bettersolutions.com/excel/formatting/number-tab-fractions.htm
}

func TestFractions(t *testing.T) {
	for _, c := range fracs {
		ff := fracFmtFunc(c.n)
		fs := ff(nil, c.v)
		if c.s != fs {
			t.Fatalf("fractions failed: got: '%s' expected: '%s' for %T(%v)",
				fs, c.s, c.v, c.v)
		}
	}
}
