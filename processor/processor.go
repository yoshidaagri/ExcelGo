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
			continue
		}

		if len(res.changes) > 0 {
			allChanges = append(allChanges, res.changes...)
			totalReplacements += len(res.changes)
		}
	}

	return totalReplacements, allChanges, nil
}
