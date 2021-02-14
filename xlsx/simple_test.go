package xlsx

import (
	"log"
	"testing"
)

func noTestOpen(t *testing.T) {
	wb, err := Open("test.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	sheets, err := wb.List()
	if err != nil {
		t.Fatal(err)
	}
	for _, s := range sheets {
		//log.Println(s)
		sheet, err := wb.Get(s)
		if err != nil {
			t.Fatal(err)
		}

		for sheet.Next() {
			sheet.Strings()
		}
	}
}

func TestOpen2(t *testing.T) {
	wb, err := Open("test2.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	sheets, err := wb.List()
	if err != nil {
		t.Fatal(err)
	}
	for _, s := range sheets {
		//log.Println(s)
		sheet, err := wb.Get(s)
		if err != nil {
			t.Fatal(err)
		}

		for sheet.Next() {
			sheet.Strings()
		}
	}
}
