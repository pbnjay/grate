package main

import (
	"bufio"
	"context"
	"flag"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"runtime/pprof"
	"strings"
	"time"

	"github.com/pbnjay/grate/xls"
)

var (
	logfile        = flag.String("l", "", "save processing logs to `filename.txt`")
	pretend        = flag.Bool("p", false, "pretend to output .tsv")
	infoFile       = flag.String("i", "results.txt", "`filename` to record stats about the process")
	removeNewlines = flag.Bool("r", true, "remove embedded tabs, newlines, and condense spaces in cell contents")
	trimSpaces     = flag.Bool("w", true, "trim whitespace from cell contents")
	skipBlanks     = flag.Bool("b", true, "discard blank rows from the output")
	cpuprofile     = flag.String("cpuprofile", "", "write cpu profile to file")
)

func main() {
	timeFormat := "2006-01-02 15:04:05"
	flag.Parse()

	if *cpuprofile != "" {
		f, err := os.Create(*cpuprofile)
		if err != nil {
			log.Fatal(err)
		}
		pprof.StartCPUProfile(f)
		defer pprof.StopCPUProfile()
	}

	if *logfile != "" {
		fo, err := os.Create(*logfile)
		if err != nil {
			log.Fatal(err)
		}
		defer fo.Close()
		log.SetOutput(fo)
	}

	fstats, err := os.OpenFile(*infoFile, os.O_CREATE|os.O_RDWR, 0644)
	if err != nil {
		log.Fatal(err)
	}
	defer fstats.Close()
	pos, err := fstats.Seek(0, io.SeekEnd)
	if err != nil {
		log.Fatal(err)
	}
	if pos == 0 {
		fmt.Fprintf(fstats, "time\tfilename\tsheet\trows\tcolumns\terrors\n")
	}
	for _, fn := range flag.Args() {
		nowFmt := time.Now().Format(timeFormat)
		results, err := processFile(fn)
		if err != nil {
			// returned errors are fatal
			fmt.Fprintf(fstats, "%s\t%s\t-\t-\t-\t%s\n", nowFmt, fn, err.Error())
			continue
		}

		for _, res := range results {
			e := "-"
			if res.Err != nil {
				e = res.Err.Error()
			}
			fmt.Fprintf(fstats, "%s\t%s\t%s\t%d\t%d\t%s\n", nowFmt, res.Filename, res.SheetName,
				res.NumRows, res.NumCols, e)
		}
	}
}

var (
	sanitize = regexp.MustCompile("[^a-zA-Z0-9]+")
	newlines = regexp.MustCompile("[ \n\r\t]+")
)

type stats struct {
	Filename  string
	SheetName string
	NumRows   int
	NumCols   int
	Err       error
}

type Flusher interface {
	Flush() error
}

func processFile(fn string) ([]stats, error) {
	log.Printf("Opening file '%s' ...", fn)
	wb, err := xls.Open(context.Background(), fn)
	if err != nil {
		return nil, err
	}

	results := []stats{}

	ext := filepath.Ext(fn)
	fn2 := filepath.Base(strings.TrimSuffix(fn, ext))

	for _, s := range wb.Sheets() {
		ps := stats{
			Filename:  fn,
			SheetName: s,
		}
		log.Printf("  Opening Sheet '%s'...", s)
		sheet, err := wb.Get(s)
		if err != nil {
			ps.Err = err
			results = append(results, ps)
			continue
		}
		if sheet.IsEmpty() {
			log.Println("    Empty sheet. Skipping.")
			results = append(results, ps)
			continue
		}
		s2 := sanitize.ReplaceAllString(s, "_")
		var w io.Writer = ioutil.Discard
		if !*pretend {
			f, err := os.Create(fn2 + "." + s2 + ".tsv")
			if err != nil {
				return nil, err
			}
			defer f.Close()
			w = bufio.NewWriter(f)
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
					if ps.NumCols < i {
						ps.NumCols = i
					}
				}
			}
			if nonblank || !*skipBlanks {
				fmt.Fprintln(w, strings.Join(row, "\t"))
				ps.NumRows++
			}
		}
		results = append(results, ps)

		if ff, ok := w.(Flusher); ok {
			ff.Flush()
		}
		if c, ok := w.(io.Closer); ok {
			c.Close()
		}
	}
	return results, nil
}
