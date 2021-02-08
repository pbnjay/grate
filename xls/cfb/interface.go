package cfb

import (
	"fmt"
	"io"
	"os"
)

// Document represents a Compound File Binary Format document.
type Document interface {
	// List the streams contained in the document.
	List() ([]string, error)

	// Open the named stream contained in the document.
	Open(name string) (io.ReadSeeker, error)
}

// Open a Compound File Binary Format document.
func Open(filename string) (Document, error) {
	d := &doc{}
	f, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	err = d.load(f)
	if err != nil {
		return nil, err
	}
	return d, nil
}

// List the streams contained in the document.
func (d *doc) List() ([]string, error) {
	var res []string
	for _, e := range d.dir {
		if e.ObjectType == typeStream {
			res = append(res, e.String())
		}
	}
	return res, nil
}

// Open the named stream contained in the document.
func (d *doc) Open(name string) (io.ReadSeeker, error) {
	for _, e := range d.dir {
		if e.String() == name && e.ObjectType == typeStream {
			if e.StreamSize < uint64(d.header.MiniStreamCutoffSize) {
				return d.getMiniStreamReader(uint32(e.StartingSectorLocation), e.StreamSize), nil
			} else if e.StreamSize != 0 {
				return d.getStreamReader(uint32(e.StartingSectorLocation), e.StreamSize), nil
			}
		}
	}
	return nil, fmt.Errorf("cfb: stream '%s' not found", name)
}
