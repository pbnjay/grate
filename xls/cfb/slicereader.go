package cfb

import (
	"errors"
	"io"
)

// SliceReader wraps a list of slices as a io.ReadSeeker that
// can transparently merge them into a single coherent stream.
type SliceReader struct {
	CSize  []int64
	Data   [][]byte
	Index  uint
	Offset uint
}

// Read implements the io.Reader interface.
func (s *SliceReader) Read(b []byte) (int, error) {
	if s.Index >= uint(len(s.Data)) {
		return 0, io.EOF
	}
	n := copy(b, s.Data[s.Index][s.Offset:])
	if n > 0 {
		s.Offset += uint(n)
		if s.Offset == uint(len(s.Data[s.Index])) {
			s.Offset = 0
			s.Index++
		}
		return n, nil
	}

	return 0, io.EOF
}

var x io.Seeker

// Seek implements the io.Seeker interface.
func (s *SliceReader) Seek(offset int64, whence int) (int64, error) {
	if len(s.CSize) != len(s.Data) {
		// calculate the cumulative block size cache
		s.CSize = make([]int64, len(s.Data))
		sz := int64(0)
		for i, d := range s.Data {
			s.CSize[i] = sz
			sz += int64(len(d))
		}
	}
	if s.Index >= uint(len(s.CSize)) {
		s.Index = uint(len(s.CSize) - 1)
		s.Offset = uint(len(s.Data[s.Index]))
	}
	// current offset in stream
	trueOffset := int64(s.Offset) + s.CSize[int(s.Index)]
	if offset == 0 && whence == io.SeekCurrent {
		// just asking for current position
		return trueOffset, nil
	}

	switch whence {
	case io.SeekStart:
		if offset < 0 {
			return -1, errors.New("xls: invalid seek offset")
		}
		s.Index = 0
		s.Offset = 0
		trueOffset = 0

	case io.SeekEnd:
		if offset > 0 {
			return -1, errors.New("xls: invalid seek offset")
		}

		s.Index = uint(len(s.Data) - 1)
		s.Offset = uint(len(s.Data[s.Index]))
		trueOffset = int64(s.Offset) + s.CSize[s.Index]

	default:
		// current position already defined
	}

	wantOffset := offset + trueOffset
	for trueOffset != wantOffset {
		loOffset := s.CSize[int(s.Index)]
		hiOffset := s.CSize[int(s.Index)] + int64(len(s.Data[s.Index]))
		if wantOffset > loOffset && wantOffset < hiOffset {
			s.Offset = uint(wantOffset - loOffset)
			return wantOffset, nil
		}

		if trueOffset > wantOffset {
			s.Index--
			s.Offset = 0
			trueOffset = s.CSize[int(s.Index)]
		} else if trueOffset < wantOffset {
			s.Index++
			s.Offset = 0
			trueOffset = s.CSize[int(s.Index)]
		}
	}
	return wantOffset, nil
}
