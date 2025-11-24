package excel

import (
	"fmt"
	"strings"

	"excel_converter/report"
	"excel_converter/utils"

	"github.com/xuri/excelize/v2"
)

// ProcessFile opens an Excel file, searches for text, replaces it, and styles the cell.
// If searchOnly is true, it only records the found text without modifying the file.
func ProcessFile(path, search, replace string, searchOnly bool) ([]report.Change, error) {
	// Use extended path for opening to support long paths
	extendedPath := utils.ToExtendedPath(path)
	f, err := excelize.OpenFile(extendedPath)
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
						if err := f.SetCellValue(sheetName, cellName, newValue); err != nil {
							changes = append(changes, report.Change{
								FilePath: path,
								Sheet:    sheetName,
								Cell:     cellName,
								OldValue: colCell,
								NewValue: newValue,
								Status:   "Failed",
								Message:  fmt.Sprintf("SetCellValue failed: %v", err),
							})
							continue
						}

						// Apply style
						f.SetCellStyle(sheetName, cellName, cellName, styleID)
						modified = true
					}

					status := "Found"
					if !searchOnly {
						status = "Success"
					}

					// Record change
					changes = append(changes, report.Change{
						FilePath: path,
						Sheet:    sheetName,
						Cell:     cellName,
						OldValue: colCell,
						NewValue: newValue, // In searchOnly, this will be same as OldValue
						Status:   status,
					})
				}
			}
		}
	}

	if modified && !searchOnly {
		fmt.Printf("[DEBUG] File %s has %d changes. Attempting to save...\n", path, len(changes))
		// Use SaveExcelSafe to handle long paths
		if err := utils.SaveExcelSafe(f, path); err != nil {
			fmt.Printf("[DEBUG] FAILED to save %s: %v\n", path, err)
			// Mark all "Success" changes as "Failed"
			for i := range changes {
				if changes[i].Status == "Success" {
					changes[i].Status = "Failed"
					changes[i].Message = fmt.Sprintf("Save failed: %v", err)
				}
			}
			// Return changes even if save failed, so they appear in the report
			return changes, fmt.Errorf("failed to save file: %w", err)
		}
		fmt.Printf("[DEBUG] Successfully saved %s\n", path)
	} else if len(changes) > 0 {
		fmt.Printf("[DEBUG] File %s has %d hits (Search Mode).\n", path, len(changes))
	}

	return changes, nil
}
