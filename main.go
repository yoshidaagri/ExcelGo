package main

import (
	"bufio"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"excel_converter/processor"
	"excel_converter/report"
	"excel_converter/server"
	"excel_converter/utils"
)

const Version = "4.7"

func main() {
	fmt.Printf("Excel Converter v%s\n", Version)

	// 1. Parse Flags
	searchFlag := flag.String("search", "", "Text to search for")
	replaceFlag := flag.String("replace", "", "Text to replace with")
	dirFlag := flag.String("dir", ".", "Directory to search in")
	serverFlag := flag.Bool("server", false, "Run in Web Server mode")
	portFlag := flag.String("port", "8080", "Port for Web Server")
	formatFlag := flag.String("format", "csv", "Output format (csv or tsv)")
	flag.Parse()

	// Check if we should run in server mode
	if *serverFlag {
		server.StartServer(*portFlag)
		return
	}

	search := *searchFlag
	replace := *replaceFlag
	rootDir := *dirFlag

	// 2. Interactive Mode if flags are missing
	reader := bufio.NewReader(os.Stdin)

	if search == "" {
		fmt.Println("Select Mode:")
		fmt.Println("1. CLI (Command Line Interface)")
		fmt.Println("2. Web GUI")
		fmt.Print("Enter choice (1 or 2) [1]: ")
		choice, _ := reader.ReadString('\n')
		choice = strings.TrimSpace(choice)

		if choice == "2" {
			server.StartServer(*portFlag)
			return
		}

		fmt.Print("検索する文字列を入力してください: ")
		input, _ := reader.ReadString('\n')
		search = strings.TrimSpace(input)
	}

	if replace == "" {
		fmt.Print("置換後の文字列を入力してください (検索モードの場合は空のままEnter): ")
		input, _ := reader.ReadString('\n')
		replace = strings.TrimSpace(input)
	}

	if search == "" {
		fmt.Println("検索文字列が指定されていません。終了します。")
		return
	}

	// Determine mode
	searchOnly := false
	if replace == "" {
		searchOnly = true
		fmt.Println("Mode: Search Only")
	} else {
		fmt.Println("Mode: Replace")
	}

	fmt.Printf("Target Directory: %s\n", rootDir)
	fmt.Printf("Search: %s\n", search)
	if !searchOnly {
		fmt.Printf("Replace: %s\n", replace)
	}
	fmt.Println("--------------------------------------------------")

	// 3. Force Close Excel
	fmt.Println("Closing Excel processes...")
	if err := utils.ForceCloseExcel(); err != nil {
		fmt.Printf("Warning: %v\n", err)
	}

	// 4. Collect Files
	fmt.Println("Scanning for Excel files...")
	files, err := processor.CollectTargetFiles(rootDir, nil, "")
	if err != nil {
		fmt.Printf("Error scanning files: %v\n", err)
		os.Exit(1)
	}
	totalFiles := len(files)
	fmt.Printf("Found %d Excel files.\n", totalFiles)
	fmt.Println("--------------------------------------------------")

	if totalFiles == 0 {
		fmt.Println("No Excel files found.")
		fmt.Println("Press Enter to exit...")
		reader.ReadString('\n')
		return
	}

	// 5. Process Files
	startTime := time.Now()
	fmt.Println("Processing files...")

	// Simple Progress Bar
	// [====================] 100% (50/50)

	totalReplacements, changes, err := processor.ProcessFiles(files, search, replace, searchOnly, func(current, total int, path string, workerCounts map[int]int) {
		percent := float64(current) / float64(total) * 100
		barLength := 50
		filledLength := int(float64(barLength) * percent / 100)
		bar := strings.Repeat("=", filledLength) + strings.Repeat(" ", barLength-filledLength)

		// \r to overwrite line
		fmt.Printf("\r[%s] %.1f%% (%d/%d) %s", bar, percent, current, total, filepath.Base(path))
		// Clear rest of line if filename is shorter than previous
		fmt.Print("                                        ")
	})
	fmt.Println() // New line after progress bar

	if err != nil {
		fmt.Printf("Error processing files: %v\n", err)
		os.Exit(1)
	}

	duration := time.Since(startTime)

	// 6. Generate Report
	if len(changes) > 0 {
		reportPath, err := report.GenerateReport(changes, rootDir, *formatFlag)
		if err != nil {
			fmt.Printf("Error generating report: %v\n", err)
		} else {
			fmt.Printf("Report generated: %s\n", reportPath)
		}
	} else {
		fmt.Println("No changes made.")
	}

	// 7. Stats
	fmt.Println("--------------------------------------------------")
	fmt.Println("Execution Summary:")
	fmt.Printf("  Time Elapsed:      %v\n", duration)
	fmt.Printf("  Files Processed:   %d\n", totalFiles)
	if searchOnly {
		fmt.Printf("  Total Hits:        %d\n", totalReplacements)
	} else {
		fmt.Printf("  Total Replacements: %d\n", totalReplacements)
	}
	fmt.Println("--------------------------------------------------")
	fmt.Println("Done.")

	// Pause to let user see output if double-clicked
	fmt.Println("Press Enter to exit...")
	reader.ReadString('\n')
}
