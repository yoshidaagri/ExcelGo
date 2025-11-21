package excel

import (
	"fmt"
	"strings"

	"excel_converter/report"

	"github.com/xuri/excelize/v2"
)

// ProcessFile opens an Excel file, searches for text, replaces it, and styles the cell.
// If searchOnly is true, it only records the found text without modifying the file.
func ProcessFile(path, search, replace string, searchOnly bool) ([]report.Change, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Printf("Error closing file %s: %v\n", path, err)
		}
	}()

	var changes []report.Change
	modified := false

	// Define the style for replaced text (Blue, Bold)
	// Only needed if not searchOnly, but defining it here is fine.
	styleID, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Color: "4180C4", // R41, G128, B196 -> Hex #2980C4 (Approx) or #4180C4?
			Bold:  true,
		},
	})
	if err != nil {
		return nil, fmt.Errorf("failed to create style: %w", err)
	}

	// Iterate over all sheets
	for _, sheetName := range f.GetSheetList() {
		rows, err := f.GetRows(sheetName)
		if err != nil {
			continue // Skip sheets we can't read
		}

		for r, row := range rows {
			for c, colCell := range row {
				if strings.Contains(colCell, search) {
					// Calculate cell name (e.g., "A1")
					cellName, _ := excelize.CoordinatesToCellName(c+1, r+1)

					newValue := colCell
					if !searchOnly {
						newValue = strings.ReplaceAll(colCell, search, replace)

						// Update cell value
						f.SetCellValue(sheetName, cellName, newValue)

						// Apply style
						f.SetCellStyle(sheetName, cellName, cellName, styleID)
						modified = true
					}

					// Record change
					changes = append(changes, report.Change{
						FilePath: path,
						Sheet:    sheetName,
						Cell:     cellName,
						OldValue: colCell,
						NewValue: newValue, // In searchOnly, this will be same as OldValue
					})
				}
			}
		}
	}

	if modified && !searchOnly {
		if err := f.Save(); err != nil {
			return nil, fmt.Errorf("failed to save file: %w", err)
		}
	}

	return changes, nil
}
