/**
 * excel-to-sql
 *
 * Excelを読み込み、ヘッダーの値をカラム名として一括インポートする
 */

package main

import (
	"errors"
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

type Params struct {
	FilePath string
	SheetNum int
	Columns  string
	Table    string
}

func main() {
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, helpMessage())
		flag.PrintDefaults()
	}

	params := &Params{}

	// 引数登録
	flag.StringVar(&params.FilePath, "file", "", "Please Input FilePath")
	flag.StringVar(&params.FilePath, "f", "", "Please Input FilePath")
	flag.IntVar(&params.SheetNum, "num", -1, "Please Input Sheet Number( 0 start )")
	flag.IntVar(&params.SheetNum, "n", -1, "Please Input Sheet Number( 0 start )")
	flag.StringVar(&params.Columns, "columns", "", "Please Input Columns")
	flag.StringVar(&params.Columns, "c", "", "Please Input Columns")
	flag.StringVar(&params.Table, "table", "", "Please Input Table")
	flag.StringVar(&params.Table, "t", "", "Please Input Table")

	flag.Parse()

	code, _ := Run(*params)
	os.Exit(code)
}

func Run(params Params) (int, error) {

	if params.FilePath == "" {
		return 1, errors.New("Error: Required FilePath")
	}

	if params.SheetNum < 0 {
		return 1, errors.New("Error: Required SheetNum")
	}

	if params.Table == "" {
		return 1, errors.New("Error: Required Table Name")
	}

	book, err := xlsx.OpenFile(params.FilePath)
	if err != nil {
		fmt.Println(err)
		return 1, errors.New("Error: Can't Open File")
	}

	var headers []string
	var outputs []string
	for s, sheet := range book.Sheets {
		if s != params.SheetNum {
			continue
		}
		for r, row := range sheet.Rows {
			if r >= 4 {
				break
			}
			var values []string
			for _, cell := range row.Cells {
				text := cell.String()
				if r == 0 {
					headers = append(headers, text)
				} else {
					values = append(values, `"`+strings.Replace(text, "\n", "\\n", -1)+`"`)
				}
			}
			if len(values) != 0 {
				outputs = append(outputs, `(`+strings.Join(values, ",")+`)`)
			}
		}
	}

	for _, output := range outputs {
		head := "INSERT INTO " + params.Table + " ( "
		if params.Columns != "" {
			head += params.Columns
		} else {
			head += strings.Join(headers, ",")
		}
		head += " ) VALUES "
		output = head + output
		fmt.Println(output + ";")
	}

	return 0, nil
}

func helpMessage() string {
	return `
Usage of excel-to-sql:

/* ja */
Excelファイル(.xlsx)を読み込んでInsert文を出力します

Options
`

}
