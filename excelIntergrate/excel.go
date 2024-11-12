package excelintergrate

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

var (
	F      *excelize.File
	axises []string
)

func Excel() {
	// Open the file
	f, err := excelize.OpenFile("C:/Users/kk/Desktop/插座购买1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	if f != nil {
		F = f
	}
}

func GetCellValByAxis(sheetName string, axis string) string {
	// Get value from cell by given worksheet name and axis.
	cell, err := F.GetCellValue(sheetName, axis)
	if err != nil {
		fmt.Println(err)
		return ""
	}
	return cell
}

func GetEachCellAxis(file *excelize.File, sheetName string, rg string) {
	axisAssemble, err := file.SearchSheet(sheetName, rg, true)
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println(axisAssemble)
	axises = axisAssemble
}

func NewIntegratedExcel() {
	// Create a new file
	f := excelize.NewFile()
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Create a new sheet.
	index, err := f.NewSheet("Sheet2")
	if err != nil {
		fmt.Println(err)
		return
	}
	// Set value of a cell.
	for _, axis := range axises {
		f.SetCellValue("Sheet2", axis, GetCellValByAxis("Sheet2", axis))
	}
	// Set active sheet of the workbook.
	f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	if err := f.SaveAs("C:/Users/kk/Desktop/ExcelIntergrate.xlsx"); err != nil {
		fmt.Println(err)
	}
}
