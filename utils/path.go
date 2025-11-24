package utils

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// ToExtendedPath converts a path to a Windows extended-length path (prefixed with \\?\).
// This allows accessing paths longer than 260 characters.
func ToExtendedPath(path string) string {
	// Clean the path first
	path = filepath.Clean(path)

	// Get absolute path
	absPath, err := filepath.Abs(path)
	if err == nil {
		path = absPath
	}

	// If already extended, return as is
	if strings.HasPrefix(path, "\\\\?\\") {
		return path
	}

	// Handle UNC paths (\\server\share -> \\?\UNC\server\share)
	if strings.HasPrefix(path, "\\\\") {
		return "\\\\?\\UNC\\" + strings.TrimPrefix(path, "\\\\")
	}

	// Regular paths (C:\foo -> \\?\C:\foo)
	return "\\\\?\\" + path
}

// SaveExcelSafe saves the excel file to a temporary location first, then moves it to the target path.
// This circumvents the 207 character limit of excelize.SaveAs by using os.Rename with an extended path.
func SaveExcelSafe(f *excelize.File, targetPath string) error {
	// 1. Create a temp file
	tempFile, err := os.CreateTemp("", "excel_converter_*.xlsx")
	if err != nil {
		return fmt.Errorf("failed to create temp file: %w", err)
	}
	tempPath := tempFile.Name()
	tempFile.Close()          // Close immediately, we just needed the name and reservation
	defer os.Remove(tempPath) // Cleanup in case of failure (Rename will make this fail harmlessly if successful)

	// 2. Save to temp file (short path)
	if err := f.SaveAs(tempPath); err != nil {
		return fmt.Errorf("failed to save to temp file: %w", err)
	}

	// 3. Move to target path using extended path
	extendedTarget := ToExtendedPath(targetPath)

	// os.Rename works for moving files on the same drive.
	// If it fails (e.g. different drive), we fall back to Copy+Delete.
	if err := os.Rename(tempPath, extendedTarget); err != nil {
		// Fallback: Copy file
		if copyErr := copyFile(tempPath, extendedTarget); copyErr != nil {
			return fmt.Errorf("failed to move file (rename: %v, copy: %v)", err, copyErr)
		}
		// If copy succeeded, remove temp file (handled by defer, but we can do it explicitly to be sure)
	}

	return nil
}

func copyFile(src, dst string) error {
	sourceFile, err := os.Open(src)
	if err != nil {
		return err
	}
	defer sourceFile.Close()

	// Ensure destination directory exists
	if err := os.MkdirAll(filepath.Dir(dst), 0755); err != nil {
		return err
	}

	destFile, err := os.Create(dst)
	if err != nil {
		return err
	}
	defer destFile.Close()

	_, err = io.Copy(destFile, sourceFile)
	return err
}
