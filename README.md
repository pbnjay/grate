# grate

A Go native tabular data extraction package. Currently supports `.xls`, `.xlsx`, `.csv`, `.tsv` formats.

# Why?

Grate focuses on speed and stability first, and makes no attempt to parse charts, figures, or other content types that may be present embedded within the input files. It tries to perform as few allocations as possible and errs on the side of caution.

There are certainly still some bugs and edge cases, but we have run it successfully on a set of 400k `.xls` and `.xlsx` files to catch many bugs and error conditions. Please file an issue with any feedback and additional problem files.

# Usage

Grate provides a simple standard interface for all supported filetypes, allowing access to both named worksheets in spreadsheets and single tables in plaintext formats.

```go
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
    wb, _ := grate.Open(os.Args[1])  // open the file
    sheets, _ := wb.List()           // list available sheets
    for _, s := range sheets {       // enumerate each sheet name
        sheet, _ := wb.Get(s)        // open the sheet
        for sheet.Next() {           // enumerate each row of data
            row := sheet.Strings()   // get the row's content as []string
            fmt.Println(strings.Join(row, "\t"))
        }
    }
    wb.Close()
}
```

# License

All source code is licensed under the [GNU GPLv3](https://raw.github.com/pbnjay/grate/master/LICENSE).
