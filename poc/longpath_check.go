package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Create a very long path
	// Max path is usually 260. We want to exceed it.
	baseDir, _ := os.Getwd()
	longDirName := strings.Repeat("a", 50)
	longPath := baseDir
	for i := 0; i < 5; i++ {
		longPath = filepath.Join(longPath, longDirName)
	}

	// Ensure directory exists
	// Note: os.MkdirAll might fail without \\?\ prefix if path is too long
	extendedPath := "\\\\?\\" + longPath
	if err := os.MkdirAll(extendedPath, 0755); err != nil {
		fmt.Printf("Failed to create directory: %v\n", err)
		return
	}
	defer os.RemoveAll("\\\\?\\" + filepath.Join(baseDir, longDirName))

	filePath := filepath.Join(longPath, "test.xlsx")
	extendedFilePath := "\\\\?\\" + filePath

	fmt.Printf("Path length: %d\n", len(filePath))
	fmt.Printf("Extended Path: %s\n", extendedFilePath)

	f := excelize.NewFile()
	// Strategy 2: Save to temp file, then rename
	tempShortPath := "temp_short.xlsx"
	if err := f.SaveAs(tempShortPath); err != nil {
		fmt.Printf("Failed to save temp file: %v\n", err)
		return
	}

	// Rename to long path
	// Note: os.Rename might need \\?\ for the destination
	if err := os.Rename(tempShortPath, extendedFilePath); err != nil {
		fmt.Printf("Failed to rename to extended path: %v\n", err)
	} else {
		fmt.Println("Successfully saved via Rename strategy!")
	}

	// Verify OpenFile
	f2, err := excelize.OpenFile(extendedFilePath)
	if err != nil {
		fmt.Printf("Failed to open extended path: %v\n", err)
	} else {
		fmt.Println("Successfully opened extended path!")
		f2.Close()
	}
}
