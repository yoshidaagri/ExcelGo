package main

import (
	"fmt"
	"log"

	"os/exec"

	"github.com/xuri/excelize/v2"
)

func main() {
	fmt.Println("Starting PoC...")

	// 1. Create a dummy Excel file
	f := excelize.NewFile()
	index, err := f.NewSheet("Sheet1")
	if err != nil {
		log.Fatal(err)
	}
	f.SetCellValue("Sheet1", "A1", "Hello World")
	f.SetCellValue("Sheet1", "B1", "This is a test")
	f.SetActiveSheet(index)

	filename := "poc_test.xlsx"
	if err := f.SaveAs(filename); err != nil {
		log.Fatal(err)
	}
	fmt.Println("Created " + filename)
	f.Close()

	// 2. Force Close Excel (Simulation)
	// In a real scenario, we would check for running processes.
	// Here we just run the command to show it works (it might fail if no excel is running, which is fine)
	cmd := exec.Command("taskkill", "/IM", "excel.exe", "/F")
	if err := cmd.Run(); err != nil {
		fmt.Println("taskkill result (expected if excel not running):", err)
	} else {
		fmt.Println("taskkill executed successfully")
	}

	// 3. Open and Modify
	f, err = excelize.OpenFile(filename)
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	// Define the style (R41, G128, B196) -> Hex #2980C4
	// 41 = 0x29, 128 = 0x80, 196 = 0xC4
	styleID, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Color: "2980C4",
			Bold:  true,
		},
	})
	if err != nil {
		log.Fatal(err)
	}

	// Search and Replace
	sheetName := "Sheet1"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	search := "World"
	replace := "Go"

	for r, row := range rows {
		for c, colCell := range row {
			if colCell == "Hello World" { // Simplified match for PoC
				cellName, _ := excelize.CoordinatesToCellName(c+1, r+1)
				fmt.Printf("Found '%s' at %s\n", search, cellName)
				
				// Replace
				newValue := "Hello " + replace
				f.SetCellValue(sheetName, cellName, newValue)
				
				// Apply Style
				err = f.SetCellStyle(sheetName, cellName, cellName, styleID)
				if err != nil {
					log.Fatal(err)
				}
				fmt.Println("Replaced and Styled")
			}
		}
	}

	if err := f.Save(); err != nil {
		log.Fatal(err)
	}
	fmt.Println("PoC Completed Successfully")
}
