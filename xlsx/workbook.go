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
	csxfNumFormat := []string{}
	d.xfs = d.xfs[:0]

	section := 0
	tok, err := dec.Token()
	for ; err == nil; tok, err = dec.Token() {
		switch v := tok.(type) {
		case xml.StartElement:
			attrs := attrMap(v.Attr)

			switch v.Name.Local {
			case "cellStyleXfs":
				section = 1
			case "cellXfs":
				section = 2
				n, _ := strconv.ParseInt(attrs["count"], 10, 64)
				d.xfs = make([]string, 0, n)

			case "xf":
				if section == 1 {
					if _, ok := attrs["applyNumberFormat"]; ok {
						csxfNumFormat = append(csxfNumFormat, attrs["numFmtId"])
					} else {
						csxfNumFormat = append(csxfNumFormat, "-")
					}
				} else if section == 2 {
					baseID, _ := strconv.ParseInt(attrs["xfId"], 10, 64)
					thisXF := csxfNumFormat[baseID]
					if _, ok := attrs["applyNumberFormat"]; ok {
						thisXF = attrs["numFmtId"]
					} else {
						thisXF = "="
					}

					nfid, _ := strconv.ParseInt(thisXF, 10, 16)
					thisXF = commonxl.BuiltInFormats[uint16(nfid)]
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
