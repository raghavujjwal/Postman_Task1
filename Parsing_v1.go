package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Student struct to hold extracted data
type Student struct {
	Emplid string  `json:"emplid"`
	Quiz   float64 `json:"quiz"`
	MidSem float64 `json:"mid_sem"`
	LabTest float64 `json:"lab_test"`
	WeeklyLabs float64 `json:"weekly_labs"`
	PreCompre float64 `json:"pre_compre"`
	Compre float64 `json:"compre"`
	Total  float64 `json:"total"`
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

func main() {
	filePath := flag.String("file", "", "Path to the Excel file")
	outputJSON := flag.String("export", "", "Export data to a JSON file")
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

	for _, row := range rows[1:] {
		if isEmptyRow(row) {
			continue
		}

		student := Student{
			Emplid:     row[columnIndices["Emplid"]],
			Quiz:      parseFloat(row[columnIndices["Quiz (30)"]]),
			MidSem:    parseFloat(row[columnIndices["Mid-Sem (75)"]]),
			LabTest:   parseFloat(row[columnIndices["Lab Test (60)"]]),
			WeeklyLabs: parseFloat(row[columnIndices["Weekly Labs (30)"]]),
			PreCompre: parseFloat(row[columnIndices["Pre-Compre (195)"]]),
			Compre:    parseFloat(row[columnIndices["Compre (105)"]]),
			Total:     parseFloat(row[columnIndices["Total (300)"]]),
		}
		students = append(students, student)
	}

	if *outputJSON != "" {
		data, err := json.MarshalIndent(students, "", "  ")
		if err != nil {
			log.Fatalln("Error generating JSON:", err)
		}
		if err := os.WriteFile(*outputJSON, data, 0644); err != nil {
			log.Fatalln("Error writing JSON file:", err)
		}
		fmt.Println("Data exported to", *outputJSON)
	} else {
		for _, student := range students {
			fmt.Printf("%s: Total Marks: %.2f\n", student.Emplid, student.Total)
		}
	}
}
