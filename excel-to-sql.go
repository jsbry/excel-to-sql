/**
 * excel-to-sql
 *
 * Excelを読み込み、ヘッダーの値をカラム名として一括インポートする
 */

package main

import (
	"bytes"
	"errors"
	"flag"
	"fmt"
	"os"
	"strings"
	"time"

	"io/ioutil"

	"github.com/cheggaaa/pb/v3"
	"github.com/tealeg/xlsx"
)

type Params struct {
	FilePath  string
	SheetNum  int
	Table     string
	Columns   string
	Output    string
	Separator int
}

func main() {
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, helpMessage())
		flag.PrintDefaults()
	}

	params := &Params{}

	// 引数登録
	flag.StringVar(&params.FilePath, "file", "", "*Required* Please Input FilePath")
	flag.StringVar(&params.FilePath, "f", "", "*Required* Please Input FilePath")
	flag.IntVar(&params.SheetNum, "num", -1, "*Required* Please Input Sheet Number( 0 start )")
	flag.IntVar(&params.SheetNum, "n", -1, "*Required* Please Input Sheet Number( 0 start )")
	flag.StringVar(&params.Table, "table", "", "*Required* Please Input Table")
	flag.StringVar(&params.Table, "t", "", "*Required* Please Input Table")
	flag.StringVar(&params.Columns, "columns", "", "Please Input Columns")
	flag.StringVar(&params.Columns, "c", "", "Please Input Columns")
	flag.StringVar(&params.Output, "output", "output.sql", "Please Input Output")
	flag.StringVar(&params.Output, "o", "output.sql", "Please Input Output")
	flag.IntVar(&params.Separator, "separator", 1, "Please Input Separator Number")
	flag.IntVar(&params.Separator, "s", 1, "Please Input Separator Number")

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

	s := time.Now()
	fmt.Println(s.Format("2006/01/02 15:04:05") + "::start")
	bar := pb.StartNew(4)

	bar.Increment()
	book, err := xlsx.OpenFile(params.FilePath)
	if err != nil {
		fmt.Println(err)
		return 1, errors.New("Error: Can't Open File")
	}
	bar.Increment()

	var headers []string
	var outputs []string
	for s, sheet := range book.Sheets {
		if s != params.SheetNum {
			continue
		}
		for r, row := range sheet.Rows {
			var values []string
			for _, cell := range row.Cells {
				text := cell.String()
				if r == 0 {
					headers = append(headers, text)
				} else {
					if text == "NULL" {
						values = append(values, `NULL`)
					} else {
						values = append(values, `"`+strings.Replace(text, "\n", "\\n", -1)+`"`)
					}
				}
			}
			if len(values) != 0 {
				outputs = append(outputs, `(`+strings.Join(values, ",")+`)`)
			}
		}
	}
	bar.Increment()

	outputsCount := len(outputs)
	var headBuffer bytes.Buffer
	var contentBuffer bytes.Buffer

	headBuffer.WriteString("INSERT INTO " + params.Table + " ( ")
	if params.Columns != "" {
		headBuffer.WriteString(params.Columns)
	} else {
		headBuffer.WriteString(strings.Join(headers, ","))
	}
	headBuffer.WriteString(" ) VALUES ")

	for o, output := range outputs {
		o++
		if o%params.Separator == 1 {
			contentBuffer.Write(headBuffer.Bytes())
		} else {

		}
		contentBuffer.WriteString(output)
		if o%params.Separator == 0 {
			contentBuffer.WriteString(";\n")
		} else {
			if outputsCount != o {
				contentBuffer.WriteString(",")
			} else {
				contentBuffer.WriteString(";\n")
			}
		}
	}
	bar.Increment()

	write_content := contentBuffer.Bytes()
	ioutil.WriteFile(params.Output, write_content, os.ModePerm)
	bar.Finish()

	f := time.Now()
	fmt.Println(f.Format("2006/01/02 15:04:05") + "::finish")

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
