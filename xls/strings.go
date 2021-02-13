package xls

import (
	"encoding/binary"
	"errors"
	"io"
	"io/ioutil"
	"unicode/utf16"
)

// 2.5.240
func decodeShortXLUnicodeString(raw []byte) (string, int, error) {
	// identical to decodeXLUnicodeString except for cch=8bits instead of 16
	cch := int(raw[0])
	flags := raw[1]
	raw = raw[2:]

	content := make([]uint16, cch)
	if (flags & 0x1) == 0 {
		// 16-bit characters but only the bottom 8bits
		contentBytes := raw[:cch]
		for i, x := range contentBytes {
			content[i] = uint16(x)
		}
		cch += 2 // to return the offset
	} else {
		// 16-bit characters
		for i := 0; i < cch; i++ {
			content[i] = binary.LittleEndian.Uint16(raw[:2])
			raw = raw[2:]
		}
		cch += cch + 2 // to return the offset
	}
	return string(utf16.Decode(content)), cch, nil
}

// 2.5.294
func decodeXLUnicodeString(raw []byte) (string, int, error) {
	// identical to decodeShortXLUnicodeString except for cch=16bits instead of 8
	cch := int(binary.LittleEndian.Uint16(raw[:2]))
	flags := raw[2]
	raw = raw[3:]

	content := make([]uint16, cch)
	if (flags & 0x1) == 0 {
		// 16-bit characters but only the bottom 8bits
		contentBytes := raw[:cch]
		for i, x := range contentBytes {
			content[i] = uint16(x)
		}
		cch += 3 // to return the offset
	} else {
		// 16-bit characters
		for i := 0; i < cch; i++ {
			content[i] = binary.LittleEndian.Uint16(raw[:2])
			raw = raw[2:]
		}
		cch += cch + 3 // to return the offset
	}
	return string(utf16.Decode(content)), cch, nil
}

// 2.5.293
func decodeXLUnicodeRichExtendedString(r io.Reader) (string, error) {
	var cch, cRun uint16
	var flags uint8
	var cbExtRs int32
	err := binary.Read(r, binary.LittleEndian, &cch)
	if err != nil {
		return "", err
	}
	err = binary.Read(r, binary.LittleEndian, &flags)
	if err != nil {
		return "", err
	}
	if (flags & 0x8) != 0 {
		// rich formating data is present
		err = binary.Read(r, binary.LittleEndian, &cRun)
		if err != nil {
			return "", err
		}
	}
	if (flags & 0x4) != 0 {
		// phonetic string data is present
		err = binary.Read(r, binary.LittleEndian, &cbExtRs)
		if err != nil {
			return "", err
		}
	}

	content := make([]uint16, cch)
	if (flags & 0x1) == 0 {
		// 16-bit characters but only the bottom 8bits
		contentBytes := make([]byte, cch)
		n, err2 := io.ReadFull(r, contentBytes)
		if n == 0 && err2 != io.ErrUnexpectedEOF {
			err = err2
		}
		if uint16(n) < cch {
			contentBytes = contentBytes[:n]
			content = content[:n]
		}

		for i, x := range contentBytes {
			content[i] = uint16(x)
		}

	} else {
		// 16-bit characters
		err = binary.Read(r, binary.LittleEndian, content)
	}
	if err != nil {
		return "", err
	}
	//////

	if cRun > 0 {
		// rich formating data is present
		_, err = io.CopyN(ioutil.Discard, r, int64(cRun)*4)
		if err != nil {
			return "", err
		}
	}
	if cbExtRs > 0 {
		// phonetic string data is present
		_, err = io.CopyN(ioutil.Discard, r, int64(cbExtRs))
		if err != nil {
			return "", err
		}
	}
	//////

	return string(utf16.Decode(content)), nil
}

// read in an array of XLUnicodeRichExtendedString s
func parseSST(recs []*rec) ([]string, error) {
	//totalRefs := binary.LittleEndian.Uint32(recs[0].Data[0:4])
	numStrings := binary.LittleEndian.Uint32(recs[0].Data[4:8])

	all := make([]string, 0, numStrings)

	buf := recs[0].Data[8:]
	for i := 0; i < len(recs); {
		var cRunBytes int
		var flags byte
		var current []uint16
		var cbExtRs uint32

		for len(buf) > 0 {
			slen := binary.LittleEndian.Uint16(buf)
			buf = buf[2:]
			flags = buf[0]
			buf = buf[1:]

			if (flags & 0x8) != 0 {
				// rich formating data is present
				cRun := binary.LittleEndian.Uint16(buf)
				cRunBytes = int(cRun) * 4
				buf = buf[2:]
			}
			if (flags & 0x4) != 0 {
				// phonetic string data is present
				cbExtRs = binary.LittleEndian.Uint32(buf)
				buf = buf[4:]
			}

			///////
			blx := len(buf)
			bly := len(buf) - 5
			if blx > 5 {
				blx = 5
			}
			if bly < 0 {
				bly = 0
			}

			// this block will read the string data, but transparently
			// handle continuing across records
			current = make([]uint16, slen)
			for j := 0; j < int(slen); j++ {
				if len(buf) == 0 {
					i++
					if (recs[i].Data[0] & 1) == 0 {
						flags &= 0xFE
					} else {
						flags |= 1
					}
					buf = recs[i].Data[1:]
				}

				if (flags & 1) == 0 { //8-bit
					current[j] = uint16(buf[0])
					buf = buf[1:]
				} else { //16-bit
					current[j] = uint16(binary.LittleEndian.Uint16(buf[:2]))
					buf = buf[2:]
					if len(buf) == 1 {
						return nil, errors.New("xls: off by one")
					}
				}
			}

			s := string(utf16.Decode(current))
			all = append(all, s)

			///////

			for cRunBytes > 0 {
				if len(buf) >= int(cRunBytes) {
					buf = buf[cRunBytes:]
					cRunBytes = 0
				} else {
					cRunBytes -= len(buf)
					i++
					buf = recs[i].Data
				}
			}

			for cbExtRs > 0 {
				if len(buf) >= int(cbExtRs) {
					buf = buf[cbExtRs:]
					cbExtRs = 0
				} else {
					cbExtRs -= uint32(len(buf))
					i++
					buf = recs[i].Data
				}
			}
		}
		i++
		if i < len(recs) {
			buf = recs[i].Data
		}
	}

	return all, nil
}
