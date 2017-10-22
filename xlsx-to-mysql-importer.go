/**
 * xlsx-to-mysql-importer
 *
 * Excelを読み込み、ヘッダーの値をカラム名として一括インポートする
 */

package main

import (
	"errors"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"strings"
)

func main() {
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, helpMessage())
		flag.PrintDefaults()
	}

	var filePath string
	var SheetNum int

	// 引数登録
	flag.StringVar(&filePath, "file", "", "Please Input FilePath")
	flag.StringVar(&filePath, "f", "", "Please Input FilePath")
	flag.IntVar(&SheetNum, "num", -1, "Please Input Sheet Number( 0 start )")
	flag.IntVar(&SheetNum, "n", -1, "Please Input Sheet Number( 0 start )")

	flag.Parse()

	code, _ := Run(filePath, SheetNum)
	os.Exit(code)
}

func Run(filePath string, SheetNum int) (int, error) {

	if filePath == "" {
		return 1, errors.New("Error: Required FilePath")
	}

	if SheetNum < 0 {
		return 1, errors.New("Error: Required SheetNum")
	}

	book, err := xlsx.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		return 1, errors.New("Error: Can't Open File")
	}

	var header, output string
	for s, sheet := range book.Sheets {
		if s != SheetNum {
			continue
		}
		for r, row := range sheet.Rows {
			if r >= 4 {
				break
			}
			var values []string
			for _, cell := range row.Cells {
				text, _ := cell.String()
				if r == 0 {
					header += text + ","
				} else {
					values = append(values, `"`+strings.Replace(text, "\n", "\\n", -1)+`"`)
				}
			}
			if len(values) != 0 {
				output += `(` + strings.Join(values, ",") + `)`
				output += "\n"
			}
		}
	}
	fmt.Println(header)
	fmt.Println(output)

	return 0, nil
}

func helpMessage() string {
	return `
Usage of xlsx-to-mysql-importer:

/* ja */
Excelファイル(.xlsx)を読み込んでDatabaseにInsertします

Options
`

}
