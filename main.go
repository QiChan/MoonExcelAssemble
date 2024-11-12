package main

import (
	D "MoonExcelAssemble/excelDevideDistrict"
	E "MoonExcelAssemble/excelIntergrate"
	"fmt"
)

func main() {
	fmt.Println("Hello World")
	if false {
		E.Excel()
		E.GetEachCellAxis(E.F, "Sheet2", "[^\\s]+")
		E.NewIntegratedExcel()
	} else {
		D.Excel()
		D.GetSheetNameList()
		D.GetAllCellSum()
		D.GetAllEachDistrictCellSum()
		D.FillFinalTable(false)
	}
}
