package cfb

import (
	"io/ioutil"
	"log"
	"os"
	"testing"
)

func TestHeader(t *testing.T) {
	d := &doc{}
	f, _ := os.Open("test.xls")
	err := d.load(f)
	if err != nil {
		t.Fatal(err)
	}
}

func TestHeader2(t *testing.T) {
	d := &doc{}
	f, _ := os.Open("test2.xls")
	err := d.load(f)
	if err != nil {
		t.Fatal(err)
	}
}

func TestHeader3(t *testing.T) {
	d := &doc{}
	f, _ := os.Open("test3.xls")
	err := d.load(f)
	if err != nil {
		t.Fatal(err)
	}
}

func TestHeader4(t *testing.T) {
	d := &doc{}
	f, _ := os.Open("test4.xls")
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
