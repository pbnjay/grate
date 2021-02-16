// Command grater extracts contents of the tabular files to stdout.
package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/pbnjay/grate"
	_ "github.com/pbnjay/grate/simple" // tsv and csv support
	_ "github.com/pbnjay/grate/xls"
	_ "github.com/pbnjay/grate/xlsx"
)

func main() {
	if len(os.Args) == 1 {
		fmt.Fprintf(os.Stderr, "USAGE: %s [file1.xls file2.xlsx file3.tsv ...]\n", os.Args[0])
		fmt.Fprintf(os.Stderr, "       Extracts contents of the tabular files to stdout\n")
		os.Exit(1)
	}
	for _, fn := range os.Args[1:] {
		wb, err := grate.Open(fn)
		if err != nil {
			fmt.Fprintln(os.Stderr, err)
			continue
		}

		sheets, err := wb.List()
		if err != nil {
			wb.Close()
			fmt.Fprintln(os.Stderr, err)
			continue
		}

		for _, s := range sheets {
			sheet, err := wb.Get(s)
			if err != nil {
				fmt.Fprintln(os.Stderr, err)
				continue
			}

			for sheet.Next() {
				row := sheet.Strings()
				fmt.Println(strings.Join(row, "\t"))
			}
		}
		wb.Close()
	}
}
