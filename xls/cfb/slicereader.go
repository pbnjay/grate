package cfb

import "io"

type sliceReader struct {
	data   [][]byte
	offset uint
}

func (s *sliceReader) Read(b []byte) (int, error) {
	var err error
	if s.offset >= uint(len(s.data)) {
		return 0, io.EOF
	}
	if len(b) < len(s.data[s.offset]) {
		return 0, io.ErrShortBuffer
	}

	n := copy(b, s.data[s.offset])
	if n == 0 {
		err = io.EOF
	}
	s.offset++
	return n, err
}
