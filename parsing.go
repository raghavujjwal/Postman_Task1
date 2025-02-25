package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func main() {
	file, err := excelize.OpenFile("CSF111_202425_01_GradeBook_stripped.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer file.Close()

	
	SheetName := "CSF111_202425_01_GradeBook"

	rows, err := file.GetRows(SheetName)
	if err != nil {
		fmt.Println(err)
		return
	}

	
	for _, row := range rows {
		fmt.Println(row)
	}
}

