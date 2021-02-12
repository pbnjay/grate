package grate

import "errors"

var (
	// configure at build time by adding go build arguments:
	//   -ldflags="-X github.com/pbnjay/grate.loglevel=debug"
	loglevel string = "warn"

	// Debug should be set to true to expose detailed logging.
	Debug bool = (loglevel == "debug")
)

// ErrInvalidScanType is returned by Scan for invalid arguments.
var ErrInvalidScanType = errors.New("grate: Scan only supports *bool, *int, *float64, *string, *time.Time arguments")

// ErrNotInFormat is used to auto-detect file types using the defined OpenFunc
// It is returned by OpenFunc when the code does not detect correct file formats.
var ErrNotInFormat = errors.New("grate: file is not in this format")

// ErrUnknownFormat is used when grate does not know how to open a file format.
var ErrUnknownFormat = errors.New("grate: file format is not known/supported")

type errx struct {
	errs []error
}

func (e errx) Error() string {
	return e.errs[0].Error()
}
func (e errx) Unwrap() error {
	if len(e.errs) > 1 {
		return e.errs[1]
	}
	return nil
}

// WrapErr wraps a set of errors.
func WrapErr(e ...error) error {
	if len(e) == 1 {
		return e[0]
	}
	return errx{errs: e}
}
