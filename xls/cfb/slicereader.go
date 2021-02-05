package cfb

import (
	"io"
)

type SliceReader struct {
	Data   [][]byte
	Index  uint
	Offset uint
}

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
