package xls

import (
	"encoding/binary"
	"io"
	"io/ioutil"
	"log"
	"unicode/utf16"
)

// 2.5.240
func decodeShortXLUnicodeString(r io.Reader) (string, error) {
	var cch, flags uint8
	err := binary.Read(r, binary.LittleEndian, &cch)
	if err != nil {
		return "", err
	}
	err = binary.Read(r, binary.LittleEndian, &flags)
	if err != nil {
		return "", err
	}

	content := make([]uint16, cch)
	if (flags & 0x1) == 0 {
		// 16-bit characters but only the bottom 8bits
		contentBytes := make([]byte, cch)
		n, err2 := io.ReadFull(r, contentBytes)
		if n == 0 && err2 != io.ErrUnexpectedEOF {
			err = err2
		}
		for i, x := range contentBytes {
			content[i] = uint16(x)
		}
	} else {
		// 16-bit characters
		err = binary.Read(r, binary.LittleEndian, content)
	}
	return string(utf16.Decode(content)), nil
}

// 2.5.293
func decodeXLUnicodeRichExtendedString(r io.Reader) (string, error) {
	var cch, cRun uint16
	var flags uint8
	var cbExtRs int32
	err := binary.Read(r, binary.LittleEndian, &cch)
	if err != nil {
		log.Println("x1", err)
		return "", err
	}
	err = binary.Read(r, binary.LittleEndian, &flags)
	if err != nil {
		log.Println("x2", err)
		return "", err
	}
	if (flags & 0x8) != 0 {
		log.Println("FORMATTING PRESENT")
		// rich formating data is present
		err = binary.Read(r, binary.LittleEndian, &cRun)
		if err != nil {
			log.Println("x3", err)
			return "", err
		}
	}
	if (flags & 0x4) != 0 {
		log.Println("PHONETIC PRESENT")
		// phonetic string data is present
		err = binary.Read(r, binary.LittleEndian, &cbExtRs)
		if err != nil {
			log.Println("x4", err)
			return "", err
		}
	}

	content := make([]uint16, cch)
	if (flags & 0x1) == 0 {
		log.Println("8BIT DATA", cch)
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
		log.Println("16BIT DATA", cch)
		// 16-bit characters
		err = binary.Read(r, binary.LittleEndian, content)
	}
	if err != nil {
		log.Println("x5", err)
	}
	//////

	if cRun > 0 {
		log.Println("READING FORMATTING DATA")
		// rich formating data is present
		_, err = io.CopyN(ioutil.Discard, r, int64(cRun)*4)
		if err != nil {
			log.Println("x6", err)
			return "", err
		}
	}
	if cbExtRs > 0 {
		log.Println("READING PHONETIC DATA")
		// phonetic string data is present
		n, err := io.CopyN(ioutil.Discard, r, int64(cbExtRs))
		if err != nil {
			log.Println("x7", n, cbExtRs, err)
			return "", err
		}
	}
	//////

	return string(utf16.Decode(content)), nil
}

// read in an array of XLUnicodeRichExtendedString s
func parseSST(recs []*rec) ([]string, error) {
	totalRefs := binary.LittleEndian.Uint32(recs[0].Data[0:4])
	numStrings := binary.LittleEndian.Uint32(recs[0].Data[4:8])

	// cell count limit is 65k x 256
	if numStrings > 65536*256 {
		log.Println("INVALID COUNTS total=", totalRefs, " -- n strings=", numStrings)
		totalRefs = 0
		numStrings = 65536 * 256
	}

	log.Println("total=", totalRefs, " -- n strings=", numStrings)
	all := make([]string, 0, numStrings)

	buf := recs[0].Data[8:]
	for i := 0; i < len(recs); {
		var blen int
		var cRunBytes int
		var flags byte
		var current []byte
		var cbExtRs uint32

		for len(buf) > 0 {
			slen := binary.LittleEndian.Uint16(buf)
			buf = buf[2:]
			flags = buf[0]
			buf = buf[1:]

			blen = int(slen)
			if (flags & 0x1) != 0 {
				// 16-bit characters
				blen = int(slen) * 2
			}

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

			// this block will read the string data, but transparently
			// handle continuing across records
			current = make([]byte, blen)
			n := copy(current, buf)
			current = current[:n]
			buf = buf[n:]
			for len(current) < blen {
				i++
				buf = recs[i].Data[1:] // skip flag TODO: verify always zero?

				n = int(blen) - len(current)
				if n > len(buf) {
					n = len(buf)
				}
				current = append(current, buf[:n]...)
				buf = buf[n:]
			}

			if (flags & 1) == 0 {
				s := string(current)
				all = append(all, s)
			} else {
				x := make([]uint16, len(current)/2)
				for y := 0; y < len(current); y += 2 {
					x[y/2] = binary.LittleEndian.Uint16(current[y : y+2])
				}
				s := string(utf16.Decode(x))
				all = append(all, s)
			}

			//log.Println(len(all), all[len(all)-1])
			for cRunBytes > 0 {
				if len(buf) >= int(cRunBytes) {
					buf = buf[cRunBytes:]
					cRunBytes = 0
				} else {
					cRunBytes -= len(buf)
					i++
					buf = recs[i].Data[1:] // skip flag TODO: verify always zero?
				}
			}

			for cbExtRs > 0 {
				if len(buf) >= int(cbExtRs) {
					buf = buf[cbExtRs:]
					cbExtRs = 0
				} else {
					cbExtRs -= uint32(len(buf))
					i++
					buf = recs[i].Data[1:] // skip flag TODO: verify always zero?
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
