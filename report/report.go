package report

import (
	"encoding/csv"
	"fmt"
	"os"
	"time"

	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
)

// Change represents a single replacement action in an Excel file.
type Change struct {
	FilePath string
	Sheet    string
	Cell     string
	OldValue string
	NewValue string
	Status   string // "Replaced", "Found", "Failed", "Skipped"
	Message  string // Error message or reason for skip
}

// GenerateCSV creates a CSV report of all changes.
func GenerateCSV(changes []Change, outputDir string) (string, error) {
	timestamp := time.Now().Format("20060102_150405")
	filename := fmt.Sprintf("replacement_report_%s.csv", timestamp)
	fullPath := filename
	if outputDir != "" {
		fullPath = outputDir + "\\" + filename
	}

	file, err := os.Create(fullPath)
	if err != nil {
		return "", err
	}
	defer file.Close()

	// Wrap file with Shift-JIS encoder
	writer := transform.NewWriter(file, japanese.ShiftJIS.NewEncoder())
	csvWriter := csv.NewWriter(writer)
	defer csvWriter.Flush()

	// Header
	header := []string{"File Path", "Sheet", "Cell", "Old Value", "New Value", "Status", "Message"}
	if err := csvWriter.Write(header); err != nil {
		return "", err
	}

	// Data
	for _, c := range changes {
		record := []string{c.FilePath, c.Sheet, c.Cell, c.OldValue, c.NewValue, c.Status, c.Message}
		if err := csvWriter.Write(record); err != nil {
			return "", err
		}
	}

	return fullPath, nil
}
