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

var requiredColumns = []string{
	"Quiz (30)", "Mid-Sem (75)", "Lab Test (60)", "Weekly Labs (30)",
	"Pre-Compre (195)", "Compre (105)", "Total (300)", "Emplid", "Class No.",
}

type Student struct {
	Emplid string  `json:"emplid"`
	Marks  float64 `json:"marks"`
	Rank   string  `json:"rank"`
}

type Report struct {
	GeneralAverages map[string]float64 `json:"general_averages"`
	BranchAverages  map[string]float64 `json:"branch_averages"`
	Discrepancies   []string           `json:"discrepancies"`
	TopStudents     map[string][]Student `json:"top_students"`
}

func parseFloat(value string) float64 {
	f, err := strconv.ParseFloat(strings.TrimSpace(value), 64)
	if err != nil {
		return 0.0
	}
	return f
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

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}

func main() {
	exportFlag := flag.Bool("export", false, "Export report as JSON")
	classFilter := flag.String("class", "", "Filter records by Class No.")
	flag.Parse()

	file, err := excelize.OpenFile("CSF111_202425_01_GradeBook_stripped.xlsx")
	if err != nil {
		log.Fatalln("Error opening file:", err)
	}
	defer file.Close()

	rows, err := file.GetRows("CSF111_202425_01_GradeBook")
	if err != nil {
		log.Fatalln("Error reading sheet:", err)
	}
	if len(rows) < 2 {
		log.Fatalln("Sheet does not contain enough data")
	}

	columnIndices := getColumnIndices(rows[0])
	totalStudents := 0
	componentSums := make(map[string]float64)
	branchTotals := make(map[string]float64)
	branchCounts := make(map[string]int)
		discrepancies := []string{}
	topStudents := make(map[string][]Student)

	for i, row := range rows[1:] {
		if isEmptyRow(row) {
			continue
		}

		if *classFilter != "" {
			if row[columnIndices["Class No."]] != *classFilter {
				continue
			}
		}

		computedTotal := parseFloat(row[columnIndices["Quiz (30)"]]) +
			parseFloat(row[columnIndices["Mid-Sem (75)"]]) +
			parseFloat(row[columnIndices["Lab Test (60)"]]) +
			parseFloat(row[columnIndices["Weekly Labs (30)"]]) +
			parseFloat(row[columnIndices["Pre-Compre (195)"]]) +
			parseFloat(row[columnIndices["Compre (105)"]])

		expectedTotal := parseFloat(row[columnIndices["Total (300)"]])
		if computedTotal != expectedTotal {
			discrepancies = append(discrepancies,
				fmt.Sprintf("Row %d: Computed Total = %.2f, Expected Total = %.2f", i+2, computedTotal, expectedTotal))
		}

		totalStudents++
		for _, component := range requiredColumns[:len(requiredColumns)-3] {
			componentSums[component] += parseFloat(row[columnIndices[component]])
		}
		branch := row[columnIndices["Class No."]]
		branchTotals[branch] += expectedTotal
		branchCounts[branch]++

		for _, component := range requiredColumns[:len(requiredColumns)-3] {
			marks := parseFloat(row[columnIndices[component]])
			student := Student{Emplid: row[columnIndices["Emplid"]], Marks: marks}
			topStudents[component] = append(topStudents[component], student)
		}
	}

	for component := range topStudents {
		students := topStudents[component]
		if len(students) > 3 {
			students = students[:3]
		}
		for rank, student := range students {
			student.Rank = fmt.Sprintf("%d", rank+1)
		}
		topStudents[component] = students
	}

	generalAverages := make(map[string]float64)
	for _, component := range requiredColumns[:len(requiredColumns)-3] {
		generalAverages[component] = componentSums[component] / float64(totalStudents)
	}

	branchAverages := make(map[string]float64)
	for branch, total := range branchTotals {
		branchAverages[branch] = total / float64(branchCounts[branch])
	}

	report := Report{
		GeneralAverages: generalAverages,
		BranchAverages:  branchAverages,
		Discrepancies:   discrepancies,
		TopStudents:     topStudents,
	}

	if *exportFlag {
		file, err := os.Create("report.json")
		if err != nil {
			log.Fatalln("Error creating JSON file:", err)
		}
		defer file.Close()
		reportJSON, _ := json.MarshalIndent(report, "", "  ")
		file.Write(reportJSON)
		fmt.Println("Report exported to report.json")
	} else {
		fmt.Println(report)
	}
}




