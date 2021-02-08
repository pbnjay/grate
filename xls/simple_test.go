package xls

import (
	"context"
	"log"
	"testing"
)

func TestHeader(t *testing.T) {
	wb, err := Open(context.Background(), "testdata/test.xls")
	if err != nil {
		t.Fatal(err)
	}
	log.Println(wb.filename)

	for _, s := range wb.Sheets() {
		log.Println(s)
		sheet, err := wb.Get(s)
		if err != nil {
			t.Fatal(err)
		}

		for sheet.Next() {
			log.Println(sheet.Strings())
		}
	}
}

func TestHeader2(t *testing.T) {
	wb, err := Open(context.Background(), "testdata/test2.xls")
	if err != nil {
		t.Fatal(err)
	}
	log.Println(wb.filename)
}

func TestHeader3(t *testing.T) {
	wb, err := Open(context.Background(), "testdata/test3.xls")
	if err != nil {
		t.Fatal(err)
	}
	log.Println(wb.filename)
}

func TestHeader4(t *testing.T) {

	wb, err := Open(context.Background(), "testdata/test4.xls")
	if err != nil {
		t.Fatal(err)
	}
	log.Println(wb.filename)
}
