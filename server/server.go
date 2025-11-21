package server

import (
	"embed"
	"encoding/json"
	"fmt"
	"io/fs"
	"log"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"excel_converter/processor"
	"excel_converter/report"
	"excel_converter/utils"
)

//go:embed static/*
var staticFiles embed.FS

type Request struct {
	Dir        string `json:"dir"`
	Search     string `json:"search"`
	Replace    string `json:"replace"`
	SearchOnly bool   `json:"searchOnly"`
}

type StatusResponse struct {
	Running           bool   `json:"running"`
	CurrentFile       string `json:"currentFile"`
	Progress          int    `json:"progress"` // 0-100
	TotalFiles        int    `json:"totalFiles"`
	ProcessedFiles    int    `json:"processedFiles"`
	TotalReplacements int    `json:"totalReplacements"`
	Message           string `json:"message"`
	ReportPath        string `json:"reportPath"`
}

var (
	currentStatus StatusResponse
	statusMutex   sync.Mutex
)

func StartServer(port string) {
	// Serve static files
	staticFS, err := fs.Sub(staticFiles, "static")
	if err != nil {
		log.Fatal(err)
	}
	http.Handle("/", http.FileServer(http.FS(staticFS)))

	// API Endpoints
	http.HandleFunc("/api/run", handleRun)
	http.HandleFunc("/api/status", handleStatus)
	http.HandleFunc("/api/browse", handleBrowse)
	http.HandleFunc("/api/download", handleDownload)
	http.HandleFunc("/api/shutdown", handleShutdown)

	fmt.Printf("Starting server at http://localhost:%s\n", port)
	if err := http.ListenAndServe(":"+port, nil); err != nil {
		log.Fatal(err)
	}
}

func handleBrowse(w http.ResponseWriter, r *http.Request) {
	// Use PowerShell to open a folder picker
	// We use a specific encoding strategy to ensure Japanese characters are preserved.
	// We write to a temporary file to avoid stdout encoding issues entirely, then read it back.
	tmpFile := filepath.Join(os.TempDir(), fmt.Sprintf("excel_converter_path_%d.txt", time.Now().UnixNano()))

	psScript := fmt.Sprintf(`
		Add-Type -AssemblyName System.Windows.Forms
		$f = New-Object System.Windows.Forms.FolderBrowserDialog
		$f.Description = "Select Target Directory"
		$f.ShowNewFolderButton = $true
		if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
			$f.SelectedPath | Out-File -FilePath "%s" -Encoding UTF8
		}
	`, tmpFile)

	cmd := exec.Command("powershell", "-NoProfile", "-Command", psScript)
	if err := cmd.Run(); err != nil {
		http.Error(w, fmt.Sprintf("PowerShell execution failed: %v", err), http.StatusInternalServerError)
		return
	}

	// Read the file back
	content, err := os.ReadFile(tmpFile)
	if err != nil {
		// If file doesn't exist, maybe user cancelled?
		if os.IsNotExist(err) {
			json.NewEncoder(w).Encode(map[string]string{"path": ""})
			return
		}
		http.Error(w, fmt.Sprintf("Failed to read path file: %v", err), http.StatusInternalServerError)
		return
	}
	defer os.Remove(tmpFile)

	// Remove BOM if present and trim whitespace
	path := string(content)
	path = strings.TrimPrefix(path, "\uFEFF")
	path = strings.TrimSpace(path)

	json.NewEncoder(w).Encode(map[string]string{"path": path})
}

func handleShutdown(w http.ResponseWriter, r *http.Request) {
	w.WriteHeader(http.StatusOK)
	w.Write([]byte("Shutting down..."))

	// Run in goroutine to allow response to be sent first
	go func() {
		time.Sleep(500 * time.Millisecond)
		os.Exit(0)
	}()
}

func handleRun(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	var req Request
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}

	statusMutex.Lock()
	if currentStatus.Running {
		statusMutex.Unlock()
		http.Error(w, "Already running", http.StatusConflict)
		return
	}
	// Reset status
	currentStatus = StatusResponse{
		Running: true,
		Message: "Scanning files...",
	}
	statusMutex.Unlock()

	go runProcessing(req)

	w.WriteHeader(http.StatusOK)
	json.NewEncoder(w).Encode(map[string]string{"status": "started"})
}

func runProcessing(req Request) {
	defer func() {
		statusMutex.Lock()
		currentStatus.Running = false
		statusMutex.Unlock()
	}()

	// 1. Collect Files
	files, err := processor.CollectTargetFiles(req.Dir)
	if err != nil {
		updateStatus(func(s *StatusResponse) {
			s.Message = fmt.Sprintf("Error collecting files: %v", err)
		})
		return
	}

	if len(files) == 0 {
		updateStatus(func(s *StatusResponse) {
			s.Message = "No Excel files found."
			s.Progress = 100
		})
		return
	}

	updateStatus(func(s *StatusResponse) {
		s.TotalFiles = len(files)
		s.Message = "Processing..."
	})

	// 2. Close Excel
	utils.ForceCloseExcel()

	// 3. Process
	replacements, changes, err := processor.ProcessFiles(files, req.Search, req.Replace, req.SearchOnly, func(current, total int, path string) {
		updateStatus(func(s *StatusResponse) {
			s.ProcessedFiles = current
			s.CurrentFile = filepath.Base(path)
			s.Progress = int(float64(current) / float64(total) * 100)
		})
	})

	if err != nil {
		updateStatus(func(s *StatusResponse) {
			s.Message = fmt.Sprintf("Error processing: %v", err)
		})
		return
	}

	// 4. Generate Report
	var reportPath string
	if len(changes) > 0 {
		reportPath, err = report.GenerateCSV(changes, req.Dir)
		if err != nil {
			updateStatus(func(s *StatusResponse) {
				s.Message = fmt.Sprintf("Error generating report: %v", err)
			})
			return
		}
	}

	updateStatus(func(s *StatusResponse) {
		s.TotalReplacements = replacements
		s.ReportPath = reportPath
		s.Message = "Completed"
		s.Progress = 100
	})
}

func updateStatus(updateFn func(*StatusResponse)) {
	statusMutex.Lock()
	defer statusMutex.Unlock()
	updateFn(&currentStatus)
}

func handleStatus(w http.ResponseWriter, r *http.Request) {
	statusMutex.Lock()
	defer statusMutex.Unlock()
	json.NewEncoder(w).Encode(currentStatus)
}

func handleDownload(w http.ResponseWriter, r *http.Request) {
	path := r.URL.Query().Get("path")
	if path == "" {
		http.Error(w, "Path required", http.StatusBadRequest)
		return
	}

	// Ensure filename has .csv extension for the browser
	filename := filepath.Base(path)
	if !strings.HasSuffix(strings.ToLower(filename), ".csv") {
		filename += ".csv"
	}

	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	http.ServeFile(w, r, path)
}
