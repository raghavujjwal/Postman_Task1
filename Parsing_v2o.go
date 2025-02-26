package main

import (
	"flag"
	"fmt"
	"log"
	"math"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)



type Student struct {
	Emplid        string  
	Quiz          float64 
	MidSem        float64 
	LabTest       float64 
	WeeklyLabs    float64 
	PreCompre     float64 
	Compre        float64 
	Total         float64 
	ComputedTotal float64 
	Discrepancy   bool    
}

var requiredColumns = []string{
	"Emplid", "Quiz (30)", "Mid-Sem (75)", "Lab Test (60)", "Weekly Labs (30)",
	"Pre-Compre (195)", "Compre (105)", "Total (300)",
}

func parseFloat(value string) float64 {
	f, err := strconv.ParseFloat(strings.TrimSpace(value), 64)
	if err != nil {
		return 0.0
	}
	return f
}

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}

func getColumnIndices(headers []string) map[string]int {
	columnIndices := make(map[string]int)
	for i, header := range headers {
		for _, col := range requiredColumns {
			if strings.Contains(header, col) {
				columnIndices[col] = i
			}
		}
	}
	return columnIndices
}

func computeTotal(student Student) float64 {
	
	computedPreCompre := student.Quiz + student.MidSem + student.LabTest + student.WeeklyLabs
	return computedPreCompre + student.Compre
}

func main() {
	filePath := flag.String("file", "", "Path to the Excel file")
	flag.Parse()

	if *filePath == "" {
		log.Fatalln("Please provide a file path using --file=<path>")
	}

	file, err := excelize.OpenFile(*filePath)
	if err != nil {
		log.Fatalln("Error opening file:", err)
	}
	defer file.Close()

	sheetName := file.GetSheetName(0)
	rows, err := file.GetRows(sheetName)
	if err != nil {
		log.Fatalln("Error reading sheet:", err)
	}

	if len(rows) < 2 {
		log.Fatalln("Sheet does not contain enough data")
	}

	columnIndices := getColumnIndices(rows[0])
	var students []Student

	const tolerance = 0.01 

	for i, row := range rows[1:] {
		if isEmptyRow(row) {
			continue
		}

		student := Student{
			Emplid:      row[columnIndices["Emplid"]],
			Quiz:       parseFloat(row[columnIndices["Quiz (30)"]]),
			MidSem:     parseFloat(row[columnIndices["Mid-Sem (75)"]]),
			LabTest:    parseFloat(row[columnIndices["Lab Test (60)"]]),
			WeeklyLabs: parseFloat(row[columnIndices["Weekly Labs (30)"]]),
			Compre:     parseFloat(row[columnIndices["Compre (105)"]]),
			Total:      parseFloat(row[columnIndices["Total (300)"]]),
		}
		student.PreCompre = student.Quiz + student.MidSem + student.LabTest + student.WeeklyLabs
		student.ComputedTotal = computeTotal(student)
		student.Discrepancy = math.Abs(student.ComputedTotal-student.Total) > tolerance

		if student.Discrepancy {
			fmt.Printf("Discrepancy in Row %d: Computed Total = %.2f, Expected Total = %.2f\n", i+2, student.ComputedTotal, student.Total)
		}

		students = append(students, student)
	}

	for _, student := range students {
		fmt.Printf("%s: Computed Total: %.2f, Expected Total: %.2f, Discrepancy: %t\n", student.Emplid, student.ComputedTotal, student.Total, student.Discrepancy)
	}
}



