package main

import (
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

func main() {
	if len(os.Args) != 2 {
		fmt.Println("Usage: go run main.go <inputExcelFile.xlsx>")
		return
	}

	inputWorkbook := os.Args[1]

	f := processWorkbook(inputWorkbook)
	if f != nil {
		// Save the modified workbook to a new file
		err := f.SaveAs("output.xlsx")
		if err != nil {
			fmt.Println("Error saving the output file:", err)
			return
		}
		fmt.Println("PivotTable created and data copied successfully")
	}
}

func processWorkbook(inputWorkbook string) *excelize.File {
	f, err := excelize.OpenFile(inputWorkbook)
	if err != nil {
		fmt.Println("Failed to open input workbook:", err)
		return nil
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Step 1: Copy data from the "Summary" sheet to "Hello" sheet
	summarySheetName := "Summary"
	helloSheetName := "Sheet1"

	// Get the rows and columns from the "Summary" sheet
	rows, err := f.GetRows(summarySheetName)
	if err != nil {
		fmt.Printf("Failed to get %s Sheet Rows: %v\n", summarySheetName, err)
		return nil
	}

	// Calculate the last row and column in the "Summary" sheet
	lastRow := len(rows)

	// Define the data range in the "Summary" sheet
	dataRange := fmt.Sprintf("%s!A1:M%d", summarySheetName, lastRow)

	// // Copy data from "Summary" to "Hello" sheet
	// err = f.SetSheetRow(helloSheetName, "A1", &rows)
	// if err != nil {
	//  fmt.Printf("Failed to set rows in %s: %v\n", helloSheetName, err)
	//  return nil
	// }

	// Copy data from "Summary" to "Hello" sheet --my
	for rowIndex, row := range rows {
		rowNumber := rowIndex + 1 // Excel row numbers are 1-based
		targetCell := fmt.Sprintf("A%d", rowNumber)
		err := f.SetSheetRow(helloSheetName, targetCell, &row)
		if err != nil {
			fmt.Printf("Failed to set row %d in %s: %v\n", rowNumber, helloSheetName, err)
			return nil
		}
	}

	if err := f.SetDefinedName(&excelize.DefinedName{
		Name:     "SourceData",
		RefersTo: "Sheet1!$A$1:$E$73",
		Comment:  "Custom defined name",
		Scope:    helloSheetName,
	}); err != nil {
		fmt.Println(err)
	}

	// Step 2: Create a PivotTable in "Hello" sheet
	pivotOptions := &excelize.PivotTableOptions{
		PivotTableRange: "Sheet1!$L$3:$U$15", // Adjust the range as needed
		DataRange:       dataRange,
		Rows: []excelize.PivotTableField{
			{Data: "Region"},
			{Data: "Dist"},
		},
		Columns: []excelize.PivotTableField{
			{Data: "Values"},
		},
		Data: []excelize.PivotTableField{
			{Data: "Sales MTD", Name: "Sum of Sales MTD", Subtotal: "Sum"},
			{Data: "Supplies Shamrock MTD", Name: "Sum of Supplies Shamrock MTD", Subtotal: "Sum"},
			{Data: "Supplies %", Name: "Average of Supplies", Subtotal: "Average"},
			{Data: "Supplies  $ Limit Full Month 0.6%", Name: "Sum of Supplies  $ Limit Full Month 0.6%", Subtotal: "Sum"},
			{Data: "Balance Avaiable", Name: "Sum of Balance Available", Subtotal: "Sum"},
			{Data: "Paper Shamrock MTD", Name: "Sum of Paper Shamrock MTD", Subtotal: "Sum"},
			{Data: "Paper% MTD", Name: "Average of Paper% MTD", Subtotal: "Average"},
			{Data: "Paper  $ Limit  Full Month 1.9%", Name: "Sum of Paper  $ Limit  Full Month 1.9%", Subtotal: "Sum"},
			{Data: "Balance Avaiable 2", Name: "Sum of Balance Avaiable 2", Subtotal: "Sum"},
		},
		RowGrandTotals: true,
		ColGrandTotals: true,
		ShowDrill:      true,
		ShowRowHeaders: true,
		ShowColHeaders: true,
		ShowLastColumn: true,
	}

	if err := f.AddPivotTable(pivotOptions); err != nil {
		fmt.Printf("Failed to create Pivot Table in %s: %v\n", helloSheetName, err)
		return nil
	}

	return f
}
