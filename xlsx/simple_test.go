package xlsx

import (
	"log"
	"testing"
)

func noTestOpen(t *testing.T) {
	_, err := Open("test.xlsx")
	if err != nil {
		log.Fatal(err)
	}
}

func TestOpen2(t *testing.T) {
	wb, err := Open("test2.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	for _, s := range wb.Sheets() {
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
