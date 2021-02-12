package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/pbnjay/grate"
	"github.com/pbnjay/grate/commonxl"
)

var _ = grate.Register("xlsx", 5, Open)

// Document contains an Office Open XML document.
type Document struct {
	filename   string
	r          *zip.Reader
	primaryDoc string

	// type => id => filename
	rels    map[string]map[string]string
	sheets  []*Sheet
	strings []string
	xfs     []commonxl.FmtFunc
	fmt     commonxl.Formatter
}

func Open(filename string) (grate.Source, error) {
	f, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	info, err := f.Stat()
	if err != nil {
		return nil, err
	}
	z, err := zip.NewReader(f, info.Size())
	if err != nil {
		return nil, grate.WrapErr(err, grate.ErrNotInFormat)
	}
	d := &Document{
		filename: filename,
		r:        z,
	}

	d.rels = make(map[string]map[string]string, 4)

	// parse the primary relationships
	dec, c, err := d.openXML("_rels/.rels")
	if err != nil {
		return nil, grate.WrapErr(err, grate.ErrNotInFormat)
	}
	err = d.parseRels(dec, "")
	c.Close()
	if err != nil {
		return nil, grate.WrapErr(err, grate.ErrNotInFormat)
	}
	if d.primaryDoc == "" {
		return nil, errors.New("xlsx: invalid document")
	}

	// parse the secondary relationships to primary doc
	base := filepath.Base(d.primaryDoc)
	sub := strings.TrimSuffix(d.primaryDoc, base)
	relfn := filepath.Join(sub, "_rels", base+".rels")
	dec, c, err = d.openXML(relfn)
	if err != nil {
		return nil, err
	}
	err = d.parseRels(dec, sub)
	c.Close()
	if err != nil {
		return nil, err
	}

	// parse the workbook structure
	dec, c, err = d.openXML(d.primaryDoc)
	if err != nil {
		return nil, err
	}
	err = d.parseWorkbook(dec)
	c.Close()
	if err != nil {
		return nil, err
	}

	styn := d.rels["http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"]
	for _, sst := range styn {
		// parse the shared string table
		dec, c, err = d.openXML(sst)
		if err != nil {
			return nil, err
		}
		err = d.parseStyles(dec)
		c.Close()
		if err != nil {
			return nil, err
		}
	}

	ssn := d.rels["http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"]
	for _, sst := range ssn {
		// parse the shared string table
		dec, c, err = d.openXML(sst)
		if err != nil {
			return nil, err
		}
		err = d.parseSharedStrings(dec)
		c.Close()
		if err != nil {
			return nil, err
		}
	}

	return d, nil
}

func (d *Document) openXML(name string) (*xml.Decoder, io.Closer, error) {
	if grate.Debug {
		log.Println("    openXML", name)
	}
	for _, zf := range d.r.File {
		if zf.Name == name {
			zfr, err := zf.Open()
			if err != nil {
				return nil, nil, err
			}
			dec := xml.NewDecoder(zfr)
			return dec, zfr, nil
		}
	}
	return nil, nil, io.EOF
}

func (d *Document) List() ([]string, error) {
	res := make([]string, 0, len(d.sheets))
	for _, s := range d.sheets {
		res = append(res, s.name)
	}
	return res, nil
}

func (d *Document) Get(sheetName string) (grate.Collection, error) {
	for _, s := range d.sheets {
		if s.name == sheetName {
			if s.err == errNotLoaded {
				s.err = s.parseSheet()
			}
			return s, s.err
		}
	}
	return nil, errors.New("xlsx: sheet not found")
}
