// package main

// import (
// 	"fmt"

// 	"github.com/xuri/excelize/v2"
// )

// func main() {

// 	f, err := excelize.OpenFile("sheet1.xlsx")

// 	if err != nil {

// 		fmt.Println(err)

// 		return
// 	}
// 	defer func() {
// 		if err := f.Close(); err != nil {
// 		}
// 		fmt.Println(err)
// 	}()

// 	SheetName := "summary"
// 	newSheet := "sheet1"
// 	if err := f.SetDefinedName(&excelize.DefinedName{

// 		Name: "SourceData",

// 		RefersTo: "summary!$A$1:$E$73",

// 		Comment: "Custom defined name",

// 		Scope: sheetName,
// 	}); err != nil {
// 		fmt.Println(err)
// 		return
// 	}
// 	rows, err := f.GetRows(SheetName)
// 	if err != nil {
// 		fmt.Printf("Failed to get %s Sheet Rows: %v\n", SheetName, err)
// 		return
// 	}
// 	lastRow := len(rows)
// 	dataRange := fmt.Sprintf("%s!A4:M%d", SheetName, lastRow)

// 	if err := f.AddPivotTable(&excelize.PivotTableOptions{
// 		DataRange:       dataRange,
// 		PivotTableRange: "Summary!$K$5:$V$15",
// 		Rows: []excelize.PivotTableField{
// 			{Data: "Region"},
// 			{Data: "Dist"},
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
// 	}); err != nil {
// 		fmt.Println(err)
// 		return
// 	}

// 	if err := f.SaveAs("new.xlsx"); err != nil {

// 		fmt.Println(err)
// 	}
// }
