package cfb

import (
	"io"
	"io/ioutil"
	"log"
	"os"
	"testing"
)

func TestHeader(t *testing.T) {
	d := &doc{}
	f, _ := os.Open("../testdata/test.xls")
	err := d.load(f)
	if err != nil {
		t.Fatal(err)
	}
}

func TestHeader2(t *testing.T) {
	d := &doc{}
	f, _ := os.Open("../testdata/test2.xls")
	err := d.load(f)
	if err != nil {
		t.Fatal(err)
	}
}

func TestHeader3(t *testing.T) {
	d := &doc{}
	f, _ := os.Open("../testdata/test3.xls")
	err := d.load(f)
	if err != nil {
		t.Fatal(err)
	}
}

func TestHeader4(t *testing.T) {
	d := &doc{}
	f, _ := os.Open("../testdata/test4.xls")
	err := d.load(f)
	if err != nil {
		t.Fatal(err)
	}

	log.Println(d.List())

	r, err := d.Open("Workbook")
	if err != nil {
		t.Fatal(err)
	}
	book, err := ioutil.ReadAll(r)
	if err != nil {
		t.Fatal(err)
	}
	log.Println(len(book))

	r, err = d.Open("\x05DocumentSummaryInformation")
	if err != nil {
		t.Fatal(err)
	}
	data, err := ioutil.ReadAll(r)
	if err != nil {
		t.Fatal(err)
	}
	log.Println(len(data))
}

var testSlices = [][]byte{
	{0, 1, 2, 3, 4, 5, 6, 7, 8, 9},
	{10, 11, 12, 13, 14, 15, 16, 17, 18, 19},
	{20, 21, 22, 23, 24, 25, 26, 27, 28, 29},
	{30, 31, 32, 33, 34, 35, 36, 37, 38, 39},
	{40, 41, 42, 43, 44, 45, 46, 47, 48, 49},
}

func TestSliceReader(t *testing.T) {
	sr := &SliceReader{
		Data: testSlices,
	}
	var uno, old [1]byte
	_, err := sr.Read(uno[:])
	for err == nil {
		old[0] = uno[0]
		_, err = sr.Read(uno[:])
		if err == nil && uno[0] != (old[0]+1) {
			log.Printf("read data out of order new=%d, old=%d", old[0], uno[0])
			t.Fail()
		}
	}
	sr.Seek(0, io.SeekStart)
	_, err = sr.Read(uno[:])
	for err == nil {
		old[0] = uno[0]
		_, err = sr.Read(uno[:])
		if err == nil && uno[0] != (old[0]+1) {
			log.Printf("read data out of order new=%d, old=%d", old[0], uno[0])
			t.Fail()
		}
	}
	sr.Seek(10, io.SeekStart)
	_, err = sr.Read(uno[:])
	if uno[0] != 10 {
		log.Printf("unexpected element %d (expected %d)", uno[0], 10)
		t.Fail()
	}
	sr.Seek(35, io.SeekStart)
	_, err = sr.Read(uno[:])
	if uno[0] != 35 {
		log.Printf("unexpected element %d (expected %d)", uno[0], 35)
		t.Fail()
	}
	sr.Seek(7, io.SeekCurrent)
	_, err = sr.Read(uno[:])
	if uno[0] != 43 {
		log.Printf("unexpected element %d (expected %d)", uno[0], 43)
		t.Fail()
	}
	sr.Seek(-9, io.SeekCurrent)
	_, err = sr.Read(uno[:])
	if uno[0] != 35 {
		log.Printf("unexpected element %d (expected %d)", uno[0], 35)
		t.Fail()
	}
}
