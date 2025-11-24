package processor

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"

	"excel_converter/excel"
	"excel_converter/report"
)

// CollectTargetFiles walks the directory and returns a list of Excel files to process.
func CollectTargetFiles(rootDir string, excludeExtensions []string, excludeDir string) ([]string, error) {
	var files []string

	// Normalize excludeDir for comparison
	if excludeDir != "" {
		excludeDir = filepath.Clean(excludeDir)
	}

	// Create a map for faster extension lookup
	excludeExtMap := make(map[string]bool)
	for _, ext := range excludeExtensions {
		excludeExtMap[strings.ToLower(ext)] = true
	}

	err := filepath.Walk(rootDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		// Check excluded directory
		if info.IsDir() {
			if excludeDir != "" && strings.HasPrefix(filepath.Clean(path), excludeDir) {
				return filepath.SkipDir
			}
			return nil
		}

		// Check extension
		ext := strings.ToLower(filepath.Ext(path))

		// Skip if extension is excluded
		if excludeExtMap[ext] {
			return nil
		}

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

type processResult struct {
	path     string
	changes  []report.Change
	err      error
	workerID int
}

// ProcessFiles processes the given list of Excel files using a worker pool.
// It accepts a callback function to report progress.
func ProcessFiles(files []string, search, replace string, searchOnly bool, onProgress func(current, total int, path string, workerCounts map[int]int)) (int, []report.Change, error) {
	totalFiles := len(files)
	if totalFiles == 0 {
		return 0, nil, nil
	}

	// Worker Pool Configuration
	numWorkers := 2
	jobs := make(chan string, totalFiles)
	results := make(chan processResult, totalFiles)

	// Start Workers
	var wg sync.WaitGroup
	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		workerID := i // Capture for closure
		go func() {
			defer wg.Done()
			for path := range jobs {
				changes, err := excel.ProcessFile(path, search, replace, searchOnly)
				results <- processResult{path: path, changes: changes, err: err, workerID: workerID}
			}
		}()
	}

	// Send Jobs
	for _, path := range files {
		jobs <- path
	}
	close(jobs)

	// Wait for workers in a separate goroutine to close results channel
	go func() {
		wg.Wait()
		close(results)
	}()

	// Collect Results
	var allChanges []report.Change
	totalReplacements := 0
	processedCount := 0
	workerCounts := make(map[int]int)

	for res := range results {
		processedCount++
		workerCounts[res.workerID]++

		if onProgress != nil {
			onProgress(processedCount, totalFiles, res.path, workerCounts)
		}

		if res.err != nil {
			fmt.Printf("\nError processing %s: %v\n", res.path, res.err)
			// Don't continue; we might have partial results (e.g. failed save)
		}

		if len(res.changes) > 0 {
			allChanges = append(allChanges, res.changes...)
			totalReplacements += len(res.changes)
		}
	}

	return totalReplacements, allChanges, nil
}
