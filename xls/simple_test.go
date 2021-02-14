package xls

import (
	"testing"
)

func TestHeader(t *testing.T) {
	wb, err := Open("../testdata/test.xls")
	if err != nil {
		t.Fatal(err)
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

	err = wb.Close()
	if err != nil {
		t.Fatal(err)
	}
}

func TestHeader2(t *testing.T) {
	wb, err := Open("../testdata/test2.xls")
	if err != nil {
		t.Fatal(err)
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

	err = wb.Close()
	if err != nil {
		t.Fatal(err)
	}
}

func TestHeader3(t *testing.T) {
	wb, err := Open("../testdata/test3.xls")
	if err != nil {
		t.Fatal(err)
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

	err = wb.Close()
	if err != nil {
		t.Fatal(err)
	}
}

func TestHeader4(t *testing.T) {
	wb, err := Open("../testdata/test4.xls")
	if err != nil {
		t.Fatal(err)
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

	err = wb.Close()
	if err != nil {
		t.Fatal(err)
	}
}
