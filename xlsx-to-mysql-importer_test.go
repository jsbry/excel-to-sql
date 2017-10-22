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
	code, err := Run("", 0)
	if code != 0 && err != nil {
		t.Log(code, err)
	} else {
		t.Error("Error Check NG")
	}
}

func Test_OkPath(t *testing.T) {
	code, err := Run("master_file.xlsx", 2)
	if code == 0 && err == nil {
		t.Log("OK")
	} else {
		t.Error(code, err)
	}
}
