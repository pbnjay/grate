package xls

import (
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"strings"
	"unicode/utf16"
)

func decodeHyperlinks(r io.Reader) (displayText, linkText string, err error) {

	var x uint64
	binary.Read(r, binary.LittleEndian, &x) // skip and discard classid
	binary.Read(r, binary.LittleEndian, &x)

	var flags, slen uint32
	binary.Read(r, binary.LittleEndian, &slen)
	if slen != 2 {
		return "", "", errors.New("xls: unknown hyperlink version")
	}

	binary.Read(r, binary.LittleEndian, &flags)
	if (flags & hlstmfHasDisplayName) != 0 {
		binary.Read(r, binary.LittleEndian, &slen)
		us := make([]uint16, slen)
		binary.Read(r, binary.LittleEndian, us)
		displayText = string(utf16.Decode(us))
	}

	if (flags & hlstmfHasFrameName) != 0 {
		// skip a HyperlinkString containing target Frame
		binary.Read(r, binary.LittleEndian, &slen)
		io.CopyN(ioutil.Discard, r, int64(slen*2))
	}

	if (flags & hlstmfHasMoniker) != 0 {
		if (flags & hlstmfMonikerSavedAsStr) != 0 {
			// read HyperlinkString containing the URL
			binary.Read(r, binary.LittleEndian, &slen)
			us := make([]uint16, slen)
			binary.Read(r, binary.LittleEndian, us)
			linkText = string(utf16.Decode(us))

		} else {
			var err error
			linkText, err = parseHyperlinkMoniker(r)
			if err != nil {
				return "", "", err
			}
		}
	}

	if (flags & hlstmfHasLocationStr) != 0 {
		binary.Read(r, binary.LittleEndian, &slen)
		us := make([]uint16, slen)
		binary.Read(r, binary.LittleEndian, us)
		linkText = string(utf16.Decode(us))
	}

	linkText = strings.Trim(linkText, " \v\f\t\r\n\x00")
	displayText = strings.Trim(displayText, " \v\f\t\r\n\x00")
	return
}

func parseHyperlinkMoniker(r io.Reader) (string, error) {
	var classid [16]byte
	n, err := r.Read(classid[:])
	if err != nil {
		return "", err
	}
	if n != 16 {
		return "", io.ErrShortBuffer
	}

	isURLMoniker := true
	isFileMoniker := true
	urlMonikerClassID := [16]byte{0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B}
	fileMonikerClassID := [16]byte{0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}
	for i, b := range classid {
		if urlMonikerClassID[i] != b {
			isURLMoniker = false
		}
		if fileMonikerClassID[i] != b {
			isFileMoniker = false
		}
	}
	if isURLMoniker {
		var length uint32
		binary.Read(r, binary.LittleEndian, &length)
		length /= 2
		buf := make([]uint16, length)
		binary.Read(r, binary.LittleEndian, &buf)
		if length > 12 && buf[length-13] == 0 {
			buf = buf[:length-12]
		}
		return string(utf16.Decode(buf)), nil
	}
	if isFileMoniker {
		var x uint16
		var length uint32
		binary.Read(r, binary.LittleEndian, &x)      //cAnti
		binary.Read(r, binary.LittleEndian, &length) //ansiLength
		buf := make([]byte, length)
		binary.Read(r, binary.LittleEndian, &buf)

		// skip 24 bytes for misc fixed properties
		io.CopyN(ioutil.Discard, r, 24)

		binary.Read(r, binary.LittleEndian, &length) // cbUnicodePathSize
		if length > 0 {
			io.CopyN(ioutil.Discard, r, 6)
			length -= 6
			buf2 := make([]uint16, (length-6)/2)
			binary.Read(r, binary.LittleEndian, &buf2)
			return string(utf16.Decode(buf2)), nil
		}

		return string(buf), nil
	}

	return "", fmt.Errorf("xls: unknown moniker classid")
}

// HLink flags
const (
	hlstmfHasMoniker          = uint32(0x001)
	hlstmfIsAbsolute          = uint32(0x002)
	hlstmfSiteGaveDisplayName = uint32(0x004)
	hlstmfHasLocationStr      = uint32(0x008)
	hlstmfHasDisplayName      = uint32(0x010)
	hlstmfHasGUID             = uint32(0x020)
	hlstmfHasCreationTime     = uint32(0x040)
	hlstmfHasFrameName        = uint32(0x080)
	hlstmfMonikerSavedAsStr   = uint32(0x100)
	hlstmfAbsFromGetdataRel   = uint32(0x200)
)
