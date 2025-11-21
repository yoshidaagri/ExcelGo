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

async function browseDir() {
    try {
        const response = await fetch('/api/browse');
        if (!response.ok) {
            throw new Error('フォルダ選択に失敗しました');
        }
        const data = await response.json();
        if (data.path) {
            document.getElementById('dir').value = data.path;
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

    if (!dir || !search) {
        alert('ディレクトリと検索文字列は必須です');
        return;
    }

    const searchOnly = mode === 'search';

    const payload = {
        dir: dir,
        search: search,
        replace: replace,
        searchOnly: searchOnly
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
