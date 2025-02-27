package main

import (
	"flag"
	"fmt"
	"log"
	"math"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Student struct to hold extracted data
type Student struct {
	CampusID      string  
	Branch        string  
	Quiz          float64 
	MidSem        float64 
	LabTest       float64 
	WeeklyLabs    float64 
	PreCompre     float64 
	Compre        float64 
	Total         float64 
	ComputedTotal float64 
	Discrepancy   bool    
	Emplid        string  
}

var requiredColumns = []string{
	"Campus ID", "Emplid", "Quiz (30)", "Mid-Sem (75)", "Lab Test (60)", "Weekly Labs (30)",
	"Pre-Compre (195)", "Compre (105)", "Total (300)",
}

func parseFloat(value string) float64 {
	f, err := strconv.ParseFloat(strings.TrimSpace(value), 64)
	if err != nil {
		return 0.0
	}
	return f
}

func extractBranch(campusID string) string {
	if len(campusID) >= 6 {
		return campusID[4:6] // Extract branch from the Campus ID
	}
	return "Unknown"
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
	classFilter := flag.String("class", "", "Filter by Class No")
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
	totalScores := make(map[string]float64)
	branchTotals := make(map[string]float64)
	branchCounts := make(map[string]int)
	topScores := make(map[string][]Student)

	const tolerance = 0.01 

	for i, row := range rows[1:] {
		if isEmptyRow(row) {
			continue
		}

		student := Student{
			CampusID:   row[columnIndices["Campus ID"]],
			Emplid:     row[columnIndices["Emplid"]],
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
		student.Branch = extractBranch(student.CampusID)

		if student.Discrepancy {
			fmt.Printf("Discrepancy in Row %d: Computed Total = %.2f, Expected Total = %.2f\n", i+2, student.ComputedTotal, student.Total)
		}

		if *classFilter == "" || *classFilter == student.CampusID {
			students = append(students, student)
			totalScores["Total"] += student.Total
			branchTotals[student.Branch] += student.Total
			branchCounts[student.Branch]++

			for _, component := range []string{"Quiz", "MidSem", "LabTest", "WeeklyLabs", "Compre", "Total"} {
				topScores[component] = append(topScores[component], student)
			}
		}
	}

	fmt.Println("\n### General Averages ###")
	for _, component := range []string{"Total"} {
		average := totalScores[component] / float64(len(students))
		fmt.Printf("%s Average: %.2f\n", component, average)
	}

	fmt.Println("\n### Branch-wise Averages (2024 Batch) ###")
	for branch, total := range branchTotals {
		if branchCounts[branch] > 0 {
			average := total / float64(branchCounts[branch])
			fmt.Printf("%s Average Total: %.2f\n", branch, average)
		}
	}

	fmt.Println("\n### Top 3 Students Per Component ###")
	for component, students := range topScores {
		sort.Slice(students, func(i, j int) bool {
			return students[i].Total > students[j].Total
		})
		fmt.Printf("\nTop 3 students for %s:\n", component)
		for rank, student := range students[:3] {
			fmt.Printf("Rank %d: Emplid: %s, Marks: %.2f\n", rank+1, student.Emplid, student.Total)
		}
	}
}









