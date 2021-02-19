// Command grater extracts contents of the tabular files to stdout.
package main

import (
	"bufio"
	"errors"
	"flag"
	"fmt"
	"log"
	"os"
	"time"

	"github.com/pbnjay/grate"
	_ "github.com/pbnjay/grate/simple" // tsv and csv support
	_ "github.com/pbnjay/grate/xls"
	_ "github.com/pbnjay/grate/xlsx"
)

func main() {
	if err := Main(); err != nil {
		log.Fatalf("%+v", err)
	}
}
func Main() error {
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, "USAGE: %s [file1.xls file2.xlsx file3.tsv ...]\n", os.Args[0])
		fmt.Fprintf(os.Stderr, "       Extracts contents of the tabular files to stdout\n")
	}
	flagDate := flag.String("date", "2006-01-02", "date format (Go) string")
	flagFloat := flag.String("float", "%g", "float format string")
	flag.Parse()
	if flag.NArg() == 0 {
		flag.Usage()
		return errors.New("a file is needed")
	}
	bw := bufio.NewWriter(os.Stdout)
	defer bw.Flush()
	for _, fn := range flag.Args() {
		wb, err := grate.Open(fn)
		if err != nil {
			return fmt.Errorf("open %q: %w", fn, err)
		}

		sheets, err := wb.List()
		if err != nil {
			wb.Close()
			return fmt.Errorf("list %q: %w", fn, err)
		}

		for _, s := range sheets {
			sheet, err := wb.Get(s)
			if err != nil {
				return fmt.Errorf("get %v: %w", s, err)
			}

			var values []interface{}
			var dests []interface{}
			for sheet.Next() {
				if values == nil {
					values = make([]interface{}, len(sheet.Strings()))
					dests = make([]interface{}, len(values))
					for i := 0; i < len(values); i++ {
						dests[i] = &values[i]
					}
				}
				for i := 0; i < len(values); i++ {
					values[i] = nil
				}
				if err = sheet.Scan(dests...); err != nil {
					return err
				}
				for i, v := range values {
					if i != 0 {
						bw.WriteByte('\t')
					}
					if v == nil {
						continue
					}
					switch x := v.(type) {
					case float64:
						v = fmt.Sprintf(*flagFloat, x)
					case time.Time:
						v = x.Format(*flagDate)
					default:
					}
					fmt.Fprintf(bw, "%v", v)
				}
				bw.WriteByte('\n')
			}
		}
		wb.Close()
	}
	return nil
}
