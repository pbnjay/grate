package xlsx

import (
	"encoding/xml"
	"errors"
	"io"
	"path/filepath"
	"strconv"

	"github.com/pbnjay/grate/commonxl"
)

func (d *Document) parseRels(dec *xml.Decoder, basedir string) error {
	tok, err := dec.Token()
	for ; err == nil; tok, err = dec.Token() {
		switch v := tok.(type) {
		// the tags we're interested in are all self-closing
		case xml.StartElement:
			switch v.Name.Local {
			case "Relationship":
				vals := make(map[string]string, 5)
				for _, a := range v.Attr {
					vals[a.Name.Local] = a.Value
				}
				if _, ok := d.rels[vals["Type"]]; !ok {
					d.rels[vals["Type"]] = make(map[string]string)
				}
				d.rels[vals["Type"]][vals["Id"]] = filepath.Join(basedir, vals["Target"])
				if vals["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" {
					d.primaryDoc = vals["Target"]
				}
			}
		}
	}
	if err == io.EOF {
		err = nil
	}
	return err
}

func (d *Document) parseWorkbook(dec *xml.Decoder) error {
	tok, err := dec.Token()
	for ; err == nil; tok, err = dec.Token() {
		switch v := tok.(type) {
		case xml.StartElement:
			//log.Println("start: ", v.Name.Local)

			switch v.Name.Local {
			case "sheet":
				vals := make(map[string]string, 5)
				for _, a := range v.Attr {
					vals[a.Name.Local] = a.Value
				}
				sheetID, ok1 := vals["id"]
				sheetName, ok2 := vals["name"]
				if !ok1 || !ok2 {
					return errors.New("xlsx: invalid sheet definition")
				}
				s := &Sheet{
					d:       d,
					relID:   sheetID,
					name:    sheetName,
					docname: d.rels["http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"][sheetID],
				}
				d.sheets = append(d.sheets, s)
			}
		case xml.EndElement:
			//log.Println("  end: ", v.Name.Local)
		default:
			//log.Printf("%T %+v", tok, tok)
		}
	}
	if err == io.EOF {
		err = nil
	}
	return err
}

func (d *Document) parseStyles(dec *xml.Decoder) error {
	baseNumFormats := []string{}
	d.xfs = d.xfs[:0]

	section := 0
	tok, err := dec.Token()
	for ; err == nil; tok, err = dec.Token() {
		switch v := tok.(type) {
		case xml.StartElement:
			attrs := attrMap(v.Attr)

			switch v.Name.Local {
			case "numFmt":
				fmtNo, _ := strconv.ParseInt(attrs["numFmtId"], 10, 16)
				d.fmt.Add(uint16(fmtNo), attrs["formatCode"])

			case "cellStyleXfs":
				section = 1
			case "cellXfs":
				section = 2
				n, _ := strconv.ParseInt(attrs["count"], 10, 64)
				d.xfs = make([]commonxl.FmtFunc, 0, n)

			case "xf":
				if section == 1 {
					// load base styles, but only save number format
					if _, ok := attrs["applyNumberFormat"]; ok {
						baseNumFormats = append(baseNumFormats, attrs["numFmtId"])
					} else {
						baseNumFormats = append(baseNumFormats, "0")
					}
				} else if section == 2 {
					// actual referencable cell styles
					// 1) get base style so we can inherit format properly
					baseID, _ := strconv.ParseInt(attrs["xfId"], 10, 64)
					numFmtID := baseNumFormats[baseID]

					// 2) check if this XF overrides the base format
					if _, ok := attrs["applyNumberFormat"]; ok {
						numFmtID = attrs["numFmtId"]
					} else {
						// remove the format (if it was inherited)
						numFmtID = "0"
					}

					nfid, _ := strconv.ParseInt(numFmtID, 10, 16)
					thisXF, ok := d.fmt.Get(uint16(nfid))
					if !ok {
						panic("numformat unknown")
					}
					d.xfs = append(d.xfs, thisXF)
				} else {
					panic("wheres is this xf??")
				}
			default:
				//log.Println("start: ", v.Name.Local, v.Attr)
			}
		case xml.EndElement:
			switch v.Name.Local {
			case "cellStyleXfs":
				section = 0
			case "cellXfs":
				section = 0
			}
			//log.Println("  end: ", v.Name.Local)
		default:
			//log.Printf("%T %+v", tok, tok)
		}
	}
	if err == io.EOF {
		err = nil
	}
	return err
}

func (d *Document) parseSharedStrings(dec *xml.Decoder) error {
	val := ""
	tok, err := dec.Token()
	for ; err == nil; tok, err = dec.Token() {
		switch v := tok.(type) {
		case xml.CharData:
			val += string(v)
		case xml.StartElement:
			switch v.Name.Local {
			case "si":
				val = ""
			default:
				//log.Println("start: ", v.Name.Local)
			}
		case xml.EndElement:
			if v.Name.Local == "si" {
				d.strings = append(d.strings, val)
				continue
			}
			//log.Println("  end: ", v.Name.Local)
		default:
			//log.Printf("%T %+v", tok, tok)
		}
	}
	if err == io.EOF {
		err = nil
	}
	return err
}
