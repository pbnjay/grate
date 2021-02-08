package main

import (
	"context"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/pbnjay/grate/xls"
)

func main() {
	//infoOnly := flag.Bool("i", false, "show info/stats ONLY")
	removeNewlines := flag.Bool("r", true, "remove embedded tabs, newlines, and condense spaces in cell contents")
	trimSpaces := flag.Bool("w", true, "trim whitespace from cell contents")
	skipBlanks := flag.Bool("b", true, "discard blank rows from the output")
	flag.Parse()

	sanitize := regexp.MustCompile("[^a-zA-Z0-9]+")
	newlines := regexp.MustCompile("[ \n\r\t]+")
	for _, fn := range flag.Args() {
		log.Printf("Opening file '%s' ...", fn)
		wb, err := xls.Open(context.Background(), fn)
		if err != nil {
			log.Println(err)
			continue
		}

		ext := filepath.Ext(fn)
		fn2 := filepath.Base(strings.TrimSuffix(fn, ext))

		for _, s := range wb.Sheets() {
			log.Printf("  Opening Sheet '%s'...", s)
			sheet, err := wb.Get(s)
			if err != nil {
				log.Println(err)
				continue
			}
			if sheet.IsEmpty() {
				log.Println("    Empty sheet. Skipping.")
				continue
			}
			s2 := sanitize.ReplaceAllString(s, "_")
			f, err := os.Create(fn2 + "." + s2 + ".tsv")
			if err != nil {
				log.Fatal(err)
			}

			for sheet.Next() {
				row := sheet.Strings()
				nonblank := false
				for i, x := range row {
					if *removeNewlines {
						x = newlines.ReplaceAllString(x, " ")
					}
					if *trimSpaces {
						x = strings.TrimSpace(x)
						row[i] = x
					}
					if x != "" {
						nonblank = true
					}
				}
				if nonblank || !*skipBlanks {
					fmt.Fprintln(f, strings.Join(row, "\t"))
				}
			}
			f.Close()
		}
	}
}
