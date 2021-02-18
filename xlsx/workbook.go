package xlsx

import (
	"encoding/xml"
	"errors"
	"io"
	"log"
	"path/filepath"
	"strconv"

	"github.com/pbnjay/grate"
	"github.com/pbnjay/grate/commonxl"
)

func (d *Document) parseRels(dec *xml.Decoder, basedir string) error {
	tok, err := dec.RawToken()
	for ; err == nil; tok, err = dec.RawToken() {
		switch v := tok.(type) {
		case xml.StartElement:
			switch v.Name.Local {
			case "Relationships":
				// container
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
			default:
				if grate.Debug {
					log.Println("      Unhandled relationship xml tag", v.Name.Local, v.Attr)
				}
			}
		case xml.EndElement:
			// not needed
		default:
			if grate.Debug {
				log.Printf("      Unhandled relationship xml tokens %T %+v", tok, tok)
			}
		}
	}
	if err == io.EOF {
		err = nil
	}
	return err
}

func (d *Document) parseWorkbook(dec *xml.Decoder) error {
	tok, err := dec.RawToken()
	for ; err == nil; tok, err = dec.RawToken() {
		switch v := tok.(type) {
		case xml.StartElement:
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
					err:     errNotLoaded,
				}
				d.sheets = append(d.sheets, s)
			case "workbook", "sheets":
				// containers
			default:
				if grate.Debug {
					log.Println("      Unhandled workbook xml tag", v.Name.Local, v.Attr)
				}
			}
		case xml.EndElement:
			// not needed
		default:
			if grate.Debug {
				log.Printf("      Unhandled workbook xml tokens %T %+v", tok, tok)
			}
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
	tok, err := dec.RawToken()
	for ; err == nil; tok, err = dec.RawToken() {
		switch v := tok.(type) {
		case xml.StartElement:
			switch v.Name.Local {
			case "styleSheet":
				// container
			case "numFmt":
				ax := getAttrs(v.Attr, "numFmtId", "formatCode")
				fmtNo, _ := strconv.ParseInt(ax[0], 10, 16)
				d.fmt.Add(uint16(fmtNo), ax[1])

			case "cellStyleXfs":
				section = 1
			case "cellXfs":
				section = 2
				ax := getAttrs(v.Attr, "count")
				n, _ := strconv.ParseInt(ax[0], 10, 64)
				d.xfs = make([]commonxl.FmtFunc, 0, n)

			case "xf":
				ax := getAttrs(v.Attr, "numFmtId", "applyNumberFormat", "xfId")
				if section == 1 {
					// load base styles, but only save number format
					if ax[1] == "0" {
						baseNumFormats = append(baseNumFormats, "0")
					} else {
						baseNumFormats = append(baseNumFormats, ax[0])
					}
				} else if section == 2 {
					// actual referencable cell styles
					// 1) get base style so we can inherit format properly
					baseID, _ := strconv.ParseInt(ax[2], 10, 64)
					numFmtID := "0"
					if len(baseNumFormats) > int(baseID) {
						numFmtID = baseNumFormats[baseID]
					}

					// 2) check if this XF overrides the base format
					if ax[1] == "0" {
						// remove the format (if it was inherited)
						numFmtID = "0"
					} else {
						numFmtID = ax[0]
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
				if grate.Debug {
					log.Println("  Unhandled style xml tag", v.Name.Local, v.Attr)
				}
			}
		case xml.EndElement:
			switch v.Name.Local {
			case "cellStyleXfs":
				section = 0
			case "cellXfs":
				section = 0
			}
		default:
			if grate.Debug {
				log.Printf("      Unhandled style xml tokens %T %+v", tok, tok)
			}
		}
	}
	if err == io.EOF {
		err = nil
	}
	return err
}

func (d *Document) parseSharedStrings(dec *xml.Decoder) error {
	val := ""
	tok, err := dec.RawToken()
	for ; err == nil; tok, err = dec.RawToken() {
		switch v := tok.(type) {
		case xml.CharData:
			val += string(v)
		case xml.StartElement:
			switch v.Name.Local {
			case "si":
				val = ""
			case "t":
				// no attributes to parse, we only want the CharData ...
			case "sst":
				// main container
			default:
				if grate.Debug {
					log.Println("  Unhandled SST xml tag", v.Name.Local, v.Attr)
				}
			}
		case xml.EndElement:
			if v.Name.Local == "si" {
				d.strings = append(d.strings, val)
				continue
			}
		default:
			if grate.Debug {
				log.Printf("    Unhandled SST xml token %T %+v", tok, tok)
			}
		}
	}
	if err == io.EOF {
		err = nil
	}
	return err
}
