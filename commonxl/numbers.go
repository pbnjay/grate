package commonxl

import (
	"math"
)

// DecimalToWholeFraction converts a floating point value into a whole
// number and fraction approximation with at most nn digits in the numerator
// and nd digits in the denominator.
func DecimalToWholeFraction(val float64, nn, nd int) (whole, num, den int) {
	num, den = DecimalToFraction(val, nn, nd)
	whole, num = num/den, num%den
	return
}

// DecimalToFraction converts a floating point value into a fraction
// approximation with at most nn digits in the numerator and nd
// digits in the denominator.
func DecimalToFraction(val float64, nn, nd int) (num, den int) {
	// http://web.archive.org/web/20111027100847/http://homepage.smc.edu/kennedy_john/DEC2FRAC.PDF
	sign := 1
	z := val
	if val < 0 {
		sign = -1
		z = -val
	}
	if nn == 0 {
		nn = 2
	}
	if nd == 0 {
		nd = 2
	}
	maxn := math.Pow(10.0, float64(nn)) // numerator with nn digits
	maxd := math.Pow(10.0, float64(nd)) // denominator with nd digits

	_, fracPart := math.Modf(val)
	if fracPart == 0.0 {
		return int(z) * sign, 1
	}
	if fracPart < 1e-9 {
		return sign, int(1e9)
	}
	if fracPart > 1e9 {
		return int(1e9) * sign, 1
	}

	diff := 1.0
	denom := 1.0
	numer := 0.0
	var lastDenom, lastNumer float64
	for diff > 1e-10 && z != math.Floor(z) {
		z = 1 / (z - math.Floor(z))
		tmp := denom
		denom = (denom * math.Floor(z)) + lastDenom
		lastDenom = tmp
		lastNumer = numer
		numer = math.Round(val * denom)
		if numer >= maxn || denom >= maxd {
			return sign * int(lastNumer), int(lastDenom)
		}
		diff = val - (numer / denom)
		if diff < 0.0 {
			diff = -diff
		}
	}
	return sign * int(numer), int(denom)
}
