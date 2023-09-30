// package main

// import (
// 	"fmt"

// 	"github.com/xuri/excelize/v2"
// )

// func main() {
// 	// Open an existing Excel workbook
// 	inputWorkbook := "paper.xlsx" // Replace with the actual file path
// 	f, err := excelize.OpenFile(inputWorkbook)
// 	if err != nil {
// 		fmt.Println("Failed to open input workbook:", err)
// 		return
// 	}
// 	defer func() {
// 		if err := f.Close(); err != nil {
// 			fmt.Println(err)
// 		}
// 	}()

// 	// Step 1: Define the data source range in the "Summary" sheet
// 	// summarySheetName := "Summary"
// 	dataRange := "summary" // Adjust the range as needed
// 	helloSheetName := "Sheet1"
// 	// Step 2: Create the PivotTable in the "Sheet1" sheet
// 	pivotOptions := &excelize.PivotTableOptions{
// 		PivotTableRange: fmt.Sprintf("%s!$A$5:$K$15", helloSheetName), // PivotTable starting cell in "Sheet1"
// 		DataRange:       dataRange,
// 		Rows: []excelize.PivotTableField{
// 			{Data: "Region"},
// 			{Data: "Dist"}, // Add row fields based on your data
// 		},
// 		Columns: []excelize.PivotTableField{},
// 		Data: []excelize.PivotTableField{
// 			{Data: "Sales MTD", Name: "Sum of Sales MTD", Subtotal: "Sum"},
// 			{Data: "Supplies Shamrock MTD", Name: "Sum of Supplies Shamrock MTD", Subtotal: "Sum"},
// 			{Data: "Supplies %", Name: "Average of Supplies", Subtotal: "Average"},
// 			{Data: "Supplies  $ Limit Full Month 0.6%", Name: "Sum of Supplies  $ Limit Full Month 0.6%", Subtotal: "Average"},
// 			{Data: "Balance Avaiable", Name: "Sum of Balance Available", Subtotal: "Sum"},
// 			{Data: "Paper Shamrock MTD", Name: "Sum of Paper Shamrock MTD", Subtotal: "Sum"},
// 			{Data: "Paper% MTD", Name: "Average of Paper% MTD", Subtotal: "Average"},
// 			{Data: "Paper  $ Limit  Full Month 1.9%", Name: "Sum of Paper  $ Limit  Full Month 1.9%", Subtotal: "Average"},
// 			{Data: "Balance Avaiable 2", Name: "Sum of Balance Avaiable 2", Subtotal: "Sum"},
// 		},
// 		RowGrandTotals: true,
// 		ColGrandTotals: true,
// 		ShowDrill:      true,
// 		ShowRowHeaders: true,
// 		ShowColHeaders: true,
// 		ShowLastColumn: true,
// 	}

// 	if err := f.AddPivotTable(pivotOptions); err != nil {
// 		fmt.Println("Failed to create Pivot Table:", err)
// 		return
// 	}

// 	// Save the modified workbook
// 	if err := f.Save(); err != nil {
// 		fmt.Println("Error saving the workbook:", err)
// 		return
// 	}

// 	fmt.Println("PivotTable created successfully in Sheet1")
// }
