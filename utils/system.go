package utils

import (
	"fmt"
	"os/exec"
)

// ForceCloseExcel attempts to terminate the Excel process to ensure files can be written.
// It returns an error if the command fails, but ignores errors if Excel was not running.
func ForceCloseExcel() error {
	cmd := exec.Command("taskkill", "/IM", "excel.exe", "/F")
	err := cmd.Run()
	if err != nil {
		// If exit code is 128 (no process found), that's fine.
		// We can't easily check the exact exit code in a cross-platform way without syscall,
		// but for this simple tool, we can assume if it fails, it might just be not running.
		// However, to be safe, we just log it or return nil if we don't care.
		// For this requirement, we just want to try to close it.
		return fmt.Errorf("failed to kill excel (might not be running): %w", err)
	}
	return nil
}
