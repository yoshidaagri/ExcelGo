let currentReportPath = "";

function toggleMode() {
    const mode = document.querySelector('input[name="mode"]:checked').value;
    const replaceGroup = document.getElementById('replace-group');
    if (mode === 'replace') {
        replaceGroup.style.display = 'block';
    } else {
        replaceGroup.style.display = 'none';
    }
}

async function browseDir(targetId = 'dir') {
    try {
        const currentPath = document.getElementById(targetId).value;
        const encodedPath = encodeURIComponent(currentPath);
        const response = await fetch(`/api/browse?path=${encodedPath}`);
        if (!response.ok) {
            throw new Error('フォルダ選択に失敗しました');
        }
        const data = await response.json();
        console.log("Browse response:", data);
        if (data.path) {
            const dirInput = document.getElementById(targetId);
            const cleanPath = data.path.trim();

            console.log("Updating input:", dirInput);
            console.log("Old value:", dirInput.value);
            console.log("New value:", cleanPath);

            dirInput.value = cleanPath;
            dirInput.setAttribute('value', cleanPath);

            // Visual feedback
            const originalBg = dirInput.style.backgroundColor;
            dirInput.style.backgroundColor = "#e8f5e9"; // Light green
            setTimeout(() => {
                dirInput.style.backgroundColor = originalBg;
            }, 500);

            console.log("Folder selected and updated:", cleanPath);
        } else {
            console.warn("No path returned from server");
        }
    } catch (error) {
        alert('エラー: ' + error.message);
    }
}

async function startProcess() {
    const dir = document.getElementById('dir').value;
    const search = document.getElementById('search').value;
    const replace = document.getElementById('replace').value;
    const mode = document.querySelector('input[name="mode"]:checked').value;
    const format = document.querySelector('input[name="format"]:checked').value;

    // Exclusion settings
    const excludeExtensions = [];
    if (document.getElementById('exclude-xlsx').checked) excludeExtensions.push('.xlsx');
    if (document.getElementById('exclude-xlsm').checked) excludeExtensions.push('.xlsm');

    const excludeDir = document.getElementById('exclude-dir').value;

    if (!dir || !search) {
        alert('ディレクトリと検索文字列は必須です');
        return;
    }

    const searchOnly = mode === 'search';

    const payload = {
        dir: dir,
        search: search,
        replace: replace,
        searchOnly: searchOnly,
        excludeExtensions: excludeExtensions,
        excludeDir: excludeDir,
        format: format
    };

    try {
        const response = await fetch('/api/run', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (response.ok) {
            document.getElementById('start-btn').disabled = true;
            document.getElementById('status-card').style.display = 'block';
            document.getElementById('download-area').style.display = 'none';
            pollStatus();
        } else {
            const err = await response.text();
            alert('開始エラー: ' + err);
        }
    } catch (error) {
        alert('通信エラー: ' + error.message);
    }
}

async function pollStatus() {
    const interval = setInterval(async () => {
        try {
            const response = await fetch('/api/status');
            const status = await response.json();

            // Update UI
            const progressBar = document.getElementById('progress-bar');
            progressBar.style.width = status.progress + '%';

            document.getElementById('status-text').textContent = status.message;
            document.getElementById('current-file').textContent = status.currentFile;
            document.getElementById('stat-files').textContent = status.processedFiles + ' / ' + status.totalFiles;

            const mode = document.querySelector('input[name="mode"]:checked').value;
            const label = mode === 'search' ? 'ヒット数' : '置換数';
            document.getElementById('stat-replacements-label').textContent = label;
            document.getElementById('stat-replacements').textContent = status.totalReplacements;

            // Update Worker Stats
            if (status.workerCounts) {
                const statsDiv = document.getElementById('worker-stats');
                const stats = Object.entries(status.workerCounts)
                    .map(([name, count]) => `${name}: ${count}件`)
                    .join(' | ');
                statsDiv.innerText = stats;
            }

            if (!status.running) {
                clearInterval(interval);
                document.getElementById('start-btn').disabled = false;
                if (status.reportPath) {
                    currentReportPath = status.reportPath;
                    document.getElementById('download-area').style.display = 'block';
                }
            }
        } catch (error) {
            console.error('Status poll error:', error);
        }
    }, 500);
}

function downloadReport() {
    if (currentReportPath) {
        window.location.href = `/api/download?path=${encodeURIComponent(currentReportPath)}`;
    }
}

async function shutdownServer() {
    if (!confirm('サーバーを停止しますか？\n停止後はブラウザを閉じてください。')) {
        return;
    }
    try {
        await fetch('/api/shutdown');
        document.body.innerHTML = '<div style="display:flex;justify-content:center;align-items:center;height:100vh;"><h1>サーバーを停止しました。ブラウザを閉じてください。</h1></div>';
    } catch (error) {
        alert('停止エラー: ' + error.message);
    }
}
