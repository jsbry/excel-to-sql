/**
 * test
 * xlsx-to-mysql-importer
 *
 * Excelを読み込み、ヘッダーの値をカラム名として一括インポートする
 */
package main

import (
	"testing"
)

func Test_NoPath(t *testing.T) {
	params := Params{
		FilePath: "",
		SheetNum: 0,
		Table:    "",
		Columns:  "",
	}
	code, err := Run(params)
	if code != 0 && err != nil {
		t.Log(code, err)
	} else {
		t.Error("Error Check NG")
	}
}

func Test_OkPath(t *testing.T) {
	params := Params{
		FilePath: "master.xlsx",
		SheetNum: 2,
		Table:    "hoge_master",
		Columns:  "col1, col2",
	}
	code, err := Run(params)
	if code == 0 && err == nil {
		t.Log("OK")
	} else {
		t.Error(code, err)
	}
}
