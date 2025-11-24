package utils

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestToExtendedPath(t *testing.T) {
	cwd, _ := os.Getwd()
	tests := []struct {
		input    string
		expected string
	}{
		{"C:\\test", "\\\\?\\C:\\test"},
		{"\\\\server\\share", "\\\\?\\UNC\\server\\share"},
		{"relative\\path", "\\\\?\\" + filepath.Join(cwd, "relative", "path")},
	}

	for _, tt := range tests {
		result := ToExtendedPath(tt.input)
		if result != tt.expected {
			t.Errorf("ToExtendedPath(%q) = %q, want %q", tt.input, result, tt.expected)
		}
	}
}

func TestSaveExcelSafe_LongPath(t *testing.T) {
	// Create a very long path
	baseDir := t.TempDir()
	longDirName := strings.Repeat("a", 50)
	longPath := baseDir
	for i := 0; i < 5; i++ {
		longPath = filepath.Join(longPath, longDirName)
	}

	// Create directory using extended path
	extendedDir := ToExtendedPath(longPath)
	if err := os.MkdirAll(extendedDir, 0755); err != nil {
		t.Fatalf("Failed to create long directory: %v", err)
	}

	filePath := filepath.Join(longPath, "test_save.xlsx")

	// Create a new file
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Test Content")

	// Try to save using SaveExcelSafe
	if err := SaveExcelSafe(f, filePath); err != nil {
		t.Fatalf("SaveExcelSafe failed: %v", err)
	}

	// Verify file exists
	extendedFilePath := ToExtendedPath(filePath)
	if _, err := os.Stat(extendedFilePath); err != nil {
		t.Errorf("File was not created at %s: %v", extendedFilePath, err)
	}

	// Verify content
	f2, err := excelize.OpenFile(extendedFilePath)
	if err != nil {
		t.Fatalf("Failed to open saved file: %v", err)
	}
	defer f2.Close()

	val, _ := f2.GetCellValue("Sheet1", "A1")
	if val != "Test Content" {
		t.Errorf("Expected content 'Test Content', got '%s'", val)
	}
}
