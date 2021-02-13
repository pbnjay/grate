package xls

import (
	"encoding/binary"
	"errors"
	"fmt"
	"strings"
	"unicode/utf16"
)

func decodeHyperlinks(raw []byte) (displayText, linkText string, err error) {
	raw = raw[16:] // skip classid
	slen := binary.LittleEndian.Uint32(raw[:4])
	if slen != 2 {
		return "", "", errors.New("xls: unknown hyperlink version")
	}

	flags := binary.LittleEndian.Uint32(raw[4:8])
	raw = raw[8:]
	if (flags & hlstmfHasDisplayName) != 0 {
		slen = binary.LittleEndian.Uint32(raw[:4])
		raw = raw[4:]
		us := make([]uint16, slen)
		for i := 0; i < int(slen); i++ {
			us[i] = binary.LittleEndian.Uint16(raw)
			raw = raw[2:]
		}
		displayText = string(utf16.Decode(us))
	}

	if (flags & hlstmfHasFrameName) != 0 {
		// skip a HyperlinkString containing target Frame
		slen = binary.LittleEndian.Uint32(raw[:4])
		raw = raw[4+(slen*2):]
	}

	if (flags & hlstmfHasMoniker) != 0 {
		if (flags & hlstmfMonikerSavedAsStr) != 0 {
			// read HyperlinkString containing the URL
			slen = binary.LittleEndian.Uint32(raw[:4])
			raw = raw[4:]
			us := make([]uint16, slen)
			for i := 0; i < int(slen); i++ {
				us[i] = binary.LittleEndian.Uint16(raw)
				raw = raw[2:]
			}
			linkText = string(utf16.Decode(us))

		} else {
			n := 0
			var err error
			linkText, n, err = parseHyperlinkMoniker(raw)
			raw = raw[n:]
			if err != nil {
				return "", "", err
			}
		}
	}

	if (flags & hlstmfHasLocationStr) != 0 {
		slen = binary.LittleEndian.Uint32(raw[:4])
		raw = raw[4:]
		us := make([]uint16, slen)
		for i := 0; i < int(slen); i++ {
			us[i] = binary.LittleEndian.Uint16(raw)
			raw = raw[2:]
		}
		linkText = string(utf16.Decode(us))
	}

	linkText = strings.Trim(linkText, " \v\f\t\r\n\x00")
	displayText = strings.Trim(displayText, " \v\f\t\r\n\x00")
	return
}

func parseHyperlinkMoniker(raw []byte) (string, int, error) {
	classid := raw[:16]
	no := 16

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
		length := binary.LittleEndian.Uint32(raw[no:])
		no += 4
		length /= 2
		buf := make([]uint16, length)
		for i := 0; i < int(length); i++ {
			buf[i] = binary.LittleEndian.Uint16(raw[no:])
			no += 2
		}
		if length > 12 && buf[length-13] == 0 {
			buf = buf[:length-12]
		}
		return string(utf16.Decode(buf)), no, nil
	}
	if isFileMoniker {
		//x := binary.LittleEndian.Uint16(raw[no:])        //cAnti
		length := binary.LittleEndian.Uint32(raw[no+2:]) //ansiLength
		no += 6
		buf := raw[no : no+int(length)]

		// skip 24 more bytes for misc fixed properties
		no += int(length) + 24

		length = binary.LittleEndian.Uint32(raw[no:]) // cbUnicodePathSize
		no += 4
		if length > 0 {
			no += 6
			length -= 6
			buf2 := make([]uint16, length/2)
			for i := 0; i < int(length/2); i++ {
				buf2[i] = binary.LittleEndian.Uint16(raw[no:])
				no += 2
			}
			return string(utf16.Decode(buf2)), no, nil
		}

		return string(buf), no, nil
	}

	return "", 0, fmt.Errorf("xls: unknown moniker classid")
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
