package excel

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func createTestExcel(t *testing.T, path string, content string) {
	f := excelize.NewFile()
	defer f.Close()

	index, err := f.NewSheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	f.SetCellValue("Sheet1", "A1", content)
	f.SetActiveSheet(index)

	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
}

func TestProcessFile_ReadOnly(t *testing.T) {
	// 1. Setup: Create a temporary Excel file
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "readonly.xlsx")
	createTestExcel(t, filePath, "OldValue")

	// 2. Make it read-only
	if err := os.Chmod(filePath, 0444); err != nil {
		t.Fatal(err)
	}
	// Ensure we restore permissions so it can be cleaned up
	defer os.Chmod(filePath, 0666)

	// 3. Execute: Try to replace "OldValue" with "NewValue"
	// We expect an error because Save() should fail.
	// CURRENT BEHAVIOR: It returns error, and changes are nil (or lost).
	// DESIRED BEHAVIOR: It returns changes with Status="Failed" and the error message.
	changes, err := ProcessFile(filePath, "OldValue", "NewValue", false)

	// 4. Verify
	// We expect an error because Save() failed.
	if err == nil {
		t.Fatal("Expected an error due to read-only file, but got nil")
	}

	// NOW: We expect changes to be returned even if Save failed.
	if len(changes) == 0 {
		t.Fatal("Expected changes to be returned even if Save failed, but got 0")
	}

	change := changes[0]
	if change.Status != "Failed" {
		t.Errorf("Expected status 'Failed', got '%s'", change.Status)
	}
	if change.Message == "" {
		t.Error("Expected error message in change record, got empty string")
	}
	t.Logf("Verified: Change returned with status '%s' and message: %s", change.Status, change.Message)
}
