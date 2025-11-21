package processor

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"excel_converter/excel"
	"excel_converter/report"
)

// CollectTargetFiles walks the directory and returns a list of Excel files to process.
func CollectTargetFiles(rootDir string) ([]string, error) {
	var files []string
	err := filepath.Walk(rootDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if info.IsDir() {
			return nil
		}
		// Check extension
		ext := strings.ToLower(filepath.Ext(path))
		if ext == ".xlsx" || ext == ".xlsm" {
			// Skip temporary files (start with ~$)
			if strings.HasPrefix(filepath.Base(path), "~$") {
				return nil
			}
			files = append(files, path)
		}
		return nil
	})
	return files, err
}

// ProcessFiles processes the given list of Excel files.
// It accepts a callback function to report progress.
func ProcessFiles(files []string, search, replace string, searchOnly bool, onProgress func(current, total int, path string)) (int, []report.Change, error) {
	var allChanges []report.Change
	totalReplacements := 0
	totalFiles := len(files)

	for i, path := range files {
		if onProgress != nil {
			onProgress(i+1, totalFiles, path)
		}

		changes, err := excel.ProcessFile(path, search, replace, searchOnly)
		if err != nil {
			fmt.Printf("\nError processing %s: %v\n", path, err)
			continue
		}

		if len(changes) > 0 {
			allChanges = append(allChanges, changes...)
			totalReplacements += len(changes)
		}
	}

	return totalReplacements, allChanges, nil
}
