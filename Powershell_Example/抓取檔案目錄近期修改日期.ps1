# 設定要搜尋的資料夾路徑
param (
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,
    
    [Parameter(Mandatory=$true)]
    [string]$BackupPath,
    
    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 100,  # 每次處理的檔案批次大小
    
    #[Parameter(Mandatory=$false)]
    #[string]$LogPath = "", # 記錄檔路徑，若為空則使用預設路徑

    [Parameter(Mandatory=$false)]
    [int]$DaysBack = 7  # 預設檢查最近7天的檔案
)

# 設定記錄檔路徑 (手動)

    $LogPath = Join-Path -Path "D:\LOG" -ChildPath "BackupVerification_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $LogPath2 = "D:\LOG"

# 記錄函數
function Write-Log {
    param($Message)
    $LogMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Add-Content -Path $LogPath -Value $LogMessage
    Write-Host $LogMessage
    
    # 將記錄添加到全域變數中
    $script:logMessages += $LogMessage
}

# 初始化日誌訊息陣列
$script:logMessages = @()

# 獲取伺服器名稱函數
function Get-ServerName {
    param([string]$Path)
    if ($Path -match "^\\\\([^\\]+)\\") {
        return $Matches[1]
    }
    return [System.Environment]::MachineName
}


# 設定檔案大小轉換函數
function Format-FileSize {
    param ([int64]$Size)
    
    switch ($Size) {
        {$_ -ge 1GB} { "{0:N2}GB" -f ($Size / 1GB); break }
        {$_ -ge 1MB} { "{0:N2}MB" -f ($Size / 1MB); break }
        {$_ -ge 1KB} { "{0:N0}KB" -f ($Size / 1KB); break }
        default { "{0}位元組" -f $Size }
    }
}

# 改進的檔案驗證函數
function Test-FileValidity {
    param (
        [string]$SourcePath,
        [string]$BackupPath
    )
    
    try {
        $sourceFile = Get-Item -Path $SourcePath
        $backupFile = Get-Item -Path $BackupPath
        
        $result = @{
            IsValid = $true
            SizeMatch = $true
            HashMatch = $true
            Reason = "正確"
            SourceSize = $sourceFile.Length
            BackupSize = $backupFile.Length
            SourceHash = $null
            BackupHash = $null
        }
        
        # 檢查檔案大小
        if ($sourceFile.Length -ne $backupFile.Length) {
            $result.IsValid = $false
            $result.SizeMatch = $false
            $result.Reason = "檔案大小不相符"
        }
        
        # 只有當檔案大小相同時才進行 MD5 檢查
        if ($result.SizeMatch) {
            $sourceHash = Get-FileHash -LiteralPath $SourcePath -Algorithm MD5
            $backupHash = Get-FileHash -LiteralPath $BackupPath -Algorithm MD5
            
            $result.SourceHash = $sourceHash.Hash
            $result.BackupHash = $backupHash.Hash
              
            if ($sourceHash.Hash -ne $backupHash.Hash) {
                $result.IsValid = $false
                $result.HashMatch = $false
                $result.Reason = if($result.SizeMatch) { "MD5不相符" } else { "檔案大小不相符且MD5不相符" }
            }
        }
        
        return $result
    }
    catch {
        return @{
            IsValid = $false
            SizeMatch = $false
            HashMatch = $false
            Reason = "檢查時發生錯誤: $($_.Exception.Message)"
            SourceSize = $null
            BackupSize = $null
            SourceHash = $null
            BackupHash = $null
        }
    }
}

# 新增 HTML 報告產生函數
function Export-HTMLReport {
    param(
        [array]$Results,
        [string]$OutputPath
    )

    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>備份驗證報告</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; border-radius: 40px;}
        th, td { max-width: 200px; text-overflow: ellipsis; overflow: hidden; border: 1px solid #ddd; padding: 8px; text-align: left; line-height: 1.5;}
        
th, td:hover {
    overflow: visible; /* 游標停駐時顯示完整內容 */
    white-space: normal; /* 恢復正常換行 */
    z-index: 10; /* 確保內容浮於上層 */
    position: relative; /* 避免影響表格布局 */
}

td {
    white-space: normal; /* 允許內容自動換行 */
    word-wrap: break-word; /* 強制長文字換行 */
    overflow-wrap: break-word; /* 支援更廣泛的瀏覽器 */
    padding: 8px;
}
        /* 改進分頁控制樣式 */
        .pagination-controls {
            margin: 10px 0;
            text-align: center;
        }
        
        .pagination-controls button, .pagination-controls input {
            padding: 8px 16px;
            margin: 0 4px;
            cursor: pointer;
            background-color: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 4px;
            transition: background-color 0.2s;
        }
        .pagination-controls button {
            cursor: pointer;
            background-color: #f8f9fa;
            transition: background-color 0.2s;
        }
        .pagination-controls button:hover:not(:disabled) {
            background-color: #e9ecef;
            border-color: #dee2e6;
        }
        
        .pagination-controls button:disabled {
            cursor: not-allowed;
            opacity: 0.5;
        }

        /* 新增排序進度指示器樣式 */
        .sort-indicator {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: rgba(0, 0, 0, 0.8);
            color: white;
            padding: 20px;
            border-radius: 8px;
            display: none;
            z-index: 1000;
        }
        
        .pagination-info {
            margin: 0 20px;
            display: inline-block;
        }


        tbody td:hover{
            background-color: bisque;
            .正確 { background-color: #dff0d8; color: #3c763d; }
            .異常 { background-color: #f2dede; color: #a94442; }
            .遺失 { background-color: #fcf8e3; color: #8a6d3b; }
            .highlight { background-color: yellow; }            
        }

        caption,caption2 {caption-side: top; text-align: left;width: 10em;
            height: 5ex;
            /*background-color: gold;*/
            /*border: 2px solid firebrick;*/
            border-radius: 10px;
            font-weight: bold;
            color: black;
            /*cursor: pointer;*/
          }

        /* 搜尋框容器樣式 */
        .search-container {
            position: sticky;
            top: 0;
            z-index: 1000;
            padding: 10px;
            box-shadow: 0px 2px 5px rgba(0, 0, 0, 0.1);
            margin-bottom: 20px;
        }
        
        /* 下拉選單 */
        .items-per-page {
            padding: 6px;
            font-size: 16px;
        }

        /* 表格滾動框架 */
        .table-container {
            height: 600px; /* 限制表格區域高度 */
            overflow-y: scroll;
            border: 1px solid #ddd;
            margin: 10px 0;
        }

        /* 搜尋框樣式 */
        #searchInput { 
            width: 100%;
            padding: 12px 40px 12px 40px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 16px;
            box-sizing: border-box;
        }
        
        /* 主搜尋圖示樣式 */
        .search-icon {
            position: absolute;
            left: 18px;
            top: 35%;
            transform: translateY(-50%);
            font-size: 18px;
            color: #666;
            cursor: pointer;
        }
        
        /* 清除搜尋框按鈕 */
        .clear-search {
            position: absolute;
            right: 18px;
            top: 30%;
            transform: translateY(-50%);
            font-size: 18px;
            color: #666;
            cursor: pointer;
            display: none; /* 初始隱藏 */
        }
        
        #searchInput:focus {
            outline: none;
            border-color: #4CAF50;
            box-shadow: 0 0 5px rgba(76,175,80,0.3);
        }
        
        /* 檔案名稱欄位中的搜尋圖示 */
        .filename-search {
            cursor: pointer;
            margin-left: 5px;
            opacity: 0.6;
        }
        
        .filename-search:hover {
            opacity: 1;
        }
        
        th { 
            white-space: nowrap; /* 禁止換行 */
            position: sticky; 
            top: 0px;
            z-index: 999;
            background-color: #f5f5f5;
            padding-right: 25px;
            user-select: none;
        }
        
        th:hover {
            background-color: #e0e0e0;
            cursor: pointer;
        }
        
        th::before,
        th::after {
            content: '';
            position: absolute;
            right: 8px;
            width: 0;
            height: 0;
            opacity: 0.3;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
        }
        /* 向上箭頭 */
        th::before {
            top: 40%;
            border-bottom: 5px solid #666;
        }
        /* 向下箭頭 */
        th::after {
            bottom: 40%;
            border-top: 5px solid #666;
        }
        /* 激活狀態的箭頭 */
        th.asc::before {
            opacity: 1;
            border-bottom-color: #2345a1;
        }
        
        th.desc::after {
            opacity: 1;
            border-top-color: #333;
        }
        
        th:hover::before,
        th:hover::after {
            opacity: 0.6;
        }
        
        .正確 { background-color: #dff0d8; color: #3c763d; }
        .異常 { background-color: #f2dede; color: #a94442; }
        .遺失 { background-color: #fcf8e3; color: #8a6d3b; }
        
        .highlight { background-color: yellow; }

        .info { 
            background-color: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 15px;
            margin-top: 20px;
            max-height: 400px;
            overflow-y: auto;
        }
        
        .info caption {
            margin: 0;
            white-space: pre-wrap;
            font-family: Consolas, monospace;
        }        

        .log-section { 
            background-color: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 15px;
            margin-top: 20px;
            max-height: 400px;
            overflow-y: auto;
        }
        
        .log-section pre {
            margin: 0;
            white-space: pre-wrap;
            font-family: Consolas, monospace;
        }
        
        /* 螢光標示樣式 */
        .highlight-row {
            animation: highlightFade 2s ease-out;
        }
        
        @keyframes highlightFade {
            0% { background-color: #fff3cd; }
            100% { background-color: transparent; }
            
        }
    </style>
</head>
<body>
    <h1>備份驗證報告</h1>


        <!-- 加入排序進度指示器 -->

    <div id="sortIndicator" class="sort-indicator">
        正在排序...
    </div>


    <div class="search-container">
        <span class="search-icon">🔍</span>
        <input type="text" id="searchInput" placeholder="搜尋檔案..." oninput="searchTable()">
        <span class="clear-search" onclick="clearSearch()">✖</span>
        <select class="items-per-page" id="itemsPerPageSelect" onchange="changeItemsPerPage()">
        <option value="20">每頁顯示 20 筆</option>
        <option value="50">每頁顯示 50 筆</option>
        <option value="100" selected>每頁顯示 100 筆</option>
    </select>
    </div>

    <!-- 改進的分頁控制 -->
    <div class="pagination-controls">
    <button onclick="firstPage()" id="firstButton" title="第一頁">⏪</button>
    <button onclick="previousPage()" id="prevButton" title="上一頁">◀</button>
    <span class="pagination-info">
        第 <span id="currentPage">1</span> 頁，共 <span id="totalPages">0</span> 頁
        （共 <span id="totalItems">0</span> 筆）
    </span>
    <button onclick="nextPage()" id="nextButton" title="下一頁">▶</button>
    <button onclick="lastPage()" id="lastButton" title="最後一頁">⏩</button>
    <input type="number" id="gotoPageInput" min="1" placeholder="跳至頁數" onchange="gotoPage()">
   
</div>
    <h2>INFO</h2>
    <div class="info">
    <caption2>執行環境: PowerShell $($PSVersionTable.PSVersion)</caption2>
    <br></br>
      <caption2>來源路徑: $SourcePath</caption2>
    <br></br>
      <caption2>備份路徑: $BackupPath</caption2>
    <br></br>
      <caption2>檢查範圍: 最近 $DaysBack 天內修改的檔案</caption2>
    <br></br>
      <caption2>執行環境: PowerShell $($PSVersionTable.PSVersion)</caption2>
    <br></br>
      <caption2>執行伺服器: $([System.Environment]::MachineName)</caption2>
    </div>
    <h3>檔案比對表</h3>
    <div class="info">
        <caption2>來源檔案數: $total </caption2>
        <br></br>
        <caption2>備份檔案數: $totalBackupFiles</caption2>
        <br></br>
        <caption2>顯示第 <span id="startIndex">0</span> 到 <span id="endIndex">0</span> 筆</caption>        
        </div>
<div class="table-container">
<table id="reportTable">
        <thead>

            <tr>
                <th onclick="sortTable(0)" style="width: 60px;">序號</th>
                <th onclick="sortTable(1)">檔案名稱</th>
                <th onclick="sortTable(2)">相對路徑</th>
                <th onclick="sortTable(3)">修改日期</th>
                <th onclick="sortTable(4)">來源目錄伺服器</th>
                <th onclick="sortTable(5)">來源大小</th>
                <th onclick="sortTable(6)">備份伺服器</th>
                <th onclick="sortTable(7)">備份大小</th>
                <th onclick="sortTable(8)">來源MD5</th>
                <th onclick="sortTable(9)">備份MD5</th>
                <th onclick="sortTable(10)">狀態</th>
                <th onclick="sortTable(11)">詳細說明</th>
            </tr>
        </thead>
        <tbody>
        $(foreach ($index in 0..($Results.Count-1)) {
            $item = $Results[$index]
            $escapedFileName = $item.檔案名稱 -replace "'", "\'"
            @"
            <tr class='$($item.狀態)'>
                <td>$($index + 1)</td>
                <td>$($item.檔案名稱) <span class='filename-search' onclick='searchForFile("$escapedFileName")'>🔍</span></td>
                <td>$($item.'相對路徑')</td>
                <td>$($item.'修改日期')</td>
                <td>$($item.'來源目錄伺服器')</td>
                <td>$($item.'來源大小')</td>
                <td>$($item.'備份伺服器')</td>
                <td>$($item.'備份大小')</td>
                <td>$($item.'來源MD5')</td>
                <td>$($item.'備份MD5')</td>
                <td>$($item.狀態)</td>
                <td>$($item.'詳細說明')</td>
            </tr>
"@
            })
        </tbody>
    </table>
    </div>
    <h2>執行記錄</h2>
    <div class="log-section">
        <pre>$(($script:logMessages | Out-String))</pre>
    </div>

    <script>

    
    // 分頁和顯示筆數控制
    let currentPage = 1;
    let itemsPerPage = 100;
    let allRows = []; // 所有表格行
    let filteredRows = []; // 搜尋後的表格行

    // 初始化
    document.addEventListener('DOMContentLoaded', () => {
        const tbody = document.querySelector('#reportTable tbody');
        allRows = Array.from(tbody.rows); // 抓取所有行
        filteredRows = [...allRows]; // 初始化為所有行
        updatePagination();
        showCurrentPage();
    });

    function changeItemsPerPage() {
        const select = document.getElementById('itemsPerPageSelect');
        itemsPerPage = parseInt(select.value, 10); // 更新每頁顯示筆數
        currentPage = 1; // 重置到第一頁
        updatePagination();
        showCurrentPage();
    }
    // 分頁控制函數
    function firstPage() {
        currentPage = 1;
        showCurrentPage();
    }

    function lastPage() {
        currentPage = Math.ceil(filteredRows.length / itemsPerPage);
        showCurrentPage();
    }

    function previousPage() {
        if (currentPage > 1) {
            currentPage--;
            showCurrentPage();
        }
    }

    function nextPage() {
        const totalPages = Math.ceil(filteredRows.length / itemsPerPage);
        if (currentPage < totalPages) {
            currentPage++;
            showCurrentPage();
        }
    }
    
    // 更新分頁資訊
    function updatePagination() {
        const totalItems = filteredRows.length; // 基於搜尋後的行數計算
        const totalPages = Math.ceil(totalItems / itemsPerPage);
    
        // 更新總頁數和當前頁數
        document.getElementById('totalPages').textContent = totalPages;
        document.getElementById('currentPage').textContent = currentPage;
    
        // 更新總筆數
        document.getElementById('totalItems').textContent = totalItems;
    
        // 計算顯示的範圍
        const startIndex = (currentPage - 1) * itemsPerPage + 1;
        const endIndex = Math.min(currentPage * itemsPerPage, totalItems);
    
        // 更新顯示範圍
        document.getElementById('startIndex').textContent = startIndex > totalItems ? 0 : startIndex;
        document.getElementById('endIndex').textContent = endIndex;
    
        // 更新按鈕狀態
        document.getElementById('firstButton').disabled = currentPage === 1;
        document.getElementById('prevButton').disabled = currentPage === 1;
        document.getElementById('nextButton').disabled = currentPage === totalPages;
        document.getElementById('lastButton').disabled = currentPage === totalPages;
    }
    
    
    // 顯示當前頁
    function showCurrentPage() {
        const tbody = document.querySelector('#reportTable tbody');
        tbody.innerHTML = ''; // 清空表格內容
        const startIndex = (currentPage - 1) * itemsPerPage;
        const endIndex = startIndex + itemsPerPage;
        for (let i = startIndex; i < endIndex && i < filteredRows.length; i++) {
            tbody.appendChild(filteredRows[i]);
        }
        updatePagination();
    }
    
    // 改進的排序函數
    function sortTable(n) {
        const sortIndicator = document.getElementById('sortIndicator');
        sortIndicator.style.display = 'block';
        
        // 使用 setTimeout 來允許 UI 更新
        setTimeout(() => {
            const table = document.getElementById("reportTable");
            const headers = table.getElementsByTagName("th");
            let dir = "asc";
            
            // 移除其他欄位的排序標記
            for (let i = 0; i < headers.length; i++) {
                if (i !== n) {
                    headers[i].classList.remove("asc", "desc");
                }
            }
            
            // 確定排序方向
            if (headers[n].classList.contains("asc")) {
                dir = "desc";
                headers[n].classList.remove("asc");
                headers[n].classList.add("desc");
            } else {
                headers[n].classList.remove("desc");
                headers[n].classList.add("asc");
            }
            
            // 對所有資料進行排序
            filteredRows.sort((a, b) => {
                const x = a.cells[n].textContent.toLowerCase();
                const y = b.cells[n].textContent.toLowerCase();
                
                // 特殊處理數字和日期
                if (n === 2) { // 修改日期欄位
                    return dir === "asc" 
                        ? new Date(x) - new Date(y)
                        : new Date(y) - new Date(x);
                } else if (n === 4 || n === 6) { // 大小欄位
                    const sizeToBytes = (str) => {
                        const num = parseFloat(str);
                        if (str.includes('GB')) return num * 1024 * 1024 * 1024;
                        if (str.includes('MB')) return num * 1024 * 1024;
                        if (str.includes('KB')) return num * 1024;
                        return num;
                    };
                    const xBytes = sizeToBytes(x);
                    const yBytes = sizeToBytes(y);
                    return dir === "asc" ? xBytes - yBytes : yBytes - xBytes;
                }
                
                // 一般文字排序
                return dir === "asc" 
                    ? (x > y ? 1 : -1)
                    : (x < y ? 1 : -1);
            });
            
            // 更新顯示
            showCurrentPage();
            sortIndicator.style.display = 'none';
        }, 0);
    }
    
    // 改進的搜尋函數
    function searchTable() {
        const filter = document.getElementById('searchInput').value.toUpperCase();
        filteredRows = allRows.filter(row =>
            Array.from(row.cells).some(cell =>
                cell.textContent.toUpperCase().includes(filter)
            )
        );
        currentPage = 1;
        showCurrentPage();
    }
    
    // 清除搜尋
    function clearSearch() {
        document.getElementById('searchInput').value = '';
        filteredRows = [...allRows];
        currentPage = 1;
        showCurrentPage();
    }
    
    // 搜尋特定檔案
    function searchForFile(filename) {
        const searchInput = document.getElementById('searchInput');
        searchInput.value = filename;
        searchInput.focus();
        document.querySelector('.clear-search').style.display = 'block';
        searchTable();
    }

    // 監聽搜尋框的輸入，控制清除按鈕的顯示
    document.getElementById('searchInput').addEventListener('input', function() {
        document.querySelector('.clear-search').style.display = 
            this.value.length > 0 ? 'block' : 'none';
    });
    function gotoPage() {
        const input = document.getElementById('gotoPageInput');
        const page = parseInt(input.value, 10);
        const totalPages = Math.ceil(filteredRows.length / itemsPerPage);

        if (page >= 1 && page <= totalPages) {
            currentPage = page;
            showCurrentPage();
        } else {
            alert('無效的頁數，請輸入 1 到 ' + totalPages + ' 之間的數字');
        }
    }
    </script>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
}
# 監控記憶體使用量的函數
function Get-MemoryUsage {
    $os = Get-WmiObject Win32_OperatingSystem
    $totalMemory = $os.TotalVisibleMemorySize  # 總記憶體（KB）
    $freeMemory = $os.FreePhysicalMemory      # 可用記憶體（KB）
    $usedMemory = $totalMemory - $freeMemory
    $memoryUsagePercent = [math]::Round(($usedMemory / $totalMemory) * 100, 2)
    
    return @{
        TotalMemory = [math]::Round($totalMemory / 1024, 2)  # 轉換為 MB
        UsedMemory = [math]::Round($usedMemory / 1024, 2)    # 轉換為 MB
        MemoryUsagePercent = $memoryUsagePercent
    }
}


# 監控 CPU 使用率的函數
function Get-CPUUsage {
    # 使用 Get-Counter 獲取 CPU 使用率
    try {
        $cpuUsage = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples.CookedValue
        return [math]::Round($cpuUsage, 2)
    }
    catch {
        Write-Log "無法取得 CPU 使用率: $($_.Exception.Message)"
        return $null
    }
}

# 監控當前處理程序記憶體使用量的函數
function Get-CurrentProcessMemory {
    $currentProcess = Get-Process -Id $PID
    $memoryMB = [math]::Round(($currentProcess.WorkingSet64 / 1MB), 2)
    return $memoryMB
}

try {
    Write-Log "開始執行備份驗證..."
    Write-Log "來源路徑: $SourcePath"
    Write-Log "備份路徑: $BackupPath"
    Write-Log "檢查範圍: 最近 $DaysBack 天內修改的檔案"
    Write-Log "執行環境: PowerShell $($PSVersionTable.PSVersion)"
    Write-Log "執行伺服器: $([System.Environment]::MachineName)"
    
    # 驗證路徑
    if (-not (Test-Path -Path $SourcePath)) {
        throw "來源資料夾路徑不存在: $SourcePath"
    }
    if (-not (Test-Path -Path $BackupPath)) {
        throw "備份資料夾路徑不存在: $BackupPath"
    }
    
    $startTime = Get-Date
    $cutoffDate = (Get-Date).AddDays(-$DaysBack)
    
    # 收集檔案清單
    Write-Log "正在收集檔案清單..."
    $sourceFiles = @(Get-ChildItem -Path $SourcePath -Recurse -File -ErrorAction SilentlyContinue | 
                    Where-Object { $_.LastWriteTime -ge $cutoffDate })
    $total = $sourceFiles.Count
    Write-Log "找到 $total 個檔案需要驗證 (修改日期在 $($cutoffDate.ToString('yyyy-MM-dd')) 之後)"
    
# 收集備份檔案清單
    Write-Log "正在收集備份目錄檔案清單..."
    $backupFiles = @(Get-ChildItem -Path $BackupPath -Recurse -File -ErrorAction SilentlyContinue | 
                    Where-Object { $_.LastWriteTime -ge $cutoffDate })
    $totalBackupFiles = $backupFiles.Count    
    
    Write-Log "找到 $total 個來源檔案 (修改日期在 $($cutoffDate.ToString('yyyy-MM-dd')) 之後)"
    Write-Log "找到 $totalBackupFiles 個備份檔案 (修改日期在 $($cutoffDate.ToString('yyyy-MM-dd')) 之後)"


    # 初始化結果陣列
    $results = @()
    $mismatchFiles = @()
    $processedCount = 0
    $lastProgressUpdate = Get-Date
    
    # 獲取來源和備份伺服器名稱
    $sourceServer = Get-ServerName -Path $SourcePath
    $backupServer = Get-ServerName -Path $BackupPath
    
    # 分批處理檔案
    for ($i = 0; $i -lt $total; $i += $BatchSize) {
        $batch = $sourceFiles | Select-Object -Skip $i -First $BatchSize
        
        foreach ($sourceFile in $batch) {
            $processedCount++
            $relativePath = $sourceFile.FullName.Substring($SourcePath.Length)
            $backupFile = Join-Path -Path $BackupPath -ChildPath $relativePath
            
            $fileInfo = [PSCustomObject]@{
                '檔案名稱' = $sourceFile.Name
                '相對路徑' = $relativePath
                '修改日期' = $sourceFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                '來源目錄伺服器' = $sourceServer
                '來源大小' = Format-FileSize -Size $sourceFile.Length
                '備份伺服器' = $backupServer
                '備份大小' = $null
                '來源MD5' = $null
                '備份MD5' = $null
                '狀態' = '處理中'
                '詳細說明' = ''
            }
            

            
            # 檢查備份檔案是否存在
            if (Test-Path -Path $backupFile) {
                $validationResult = Test-FileValidity -SourcePath $sourceFile.FullName -BackupPath $backupFile
                
                $fileInfo.'備份大小' = Format-FileSize -Size $validationResult.BackupSize
                $fileInfo.'來源MD5' = $validationResult.SourceHash
                $fileInfo.'備份MD5' = $validationResult.BackupHash
                
                # 新的狀態判斷
                if ($validationResult.SizeMatch -and $validationResult.HashMatch) {
                    $fileInfo.'狀態' = '正確'
                } elseif (-not $validationResult.SizeMatch -and -not $validationResult.HashMatch) {
                    $fileInfo.'狀態' = '異常'
                } elseif (-not $validationResult.SizeMatch) {
                    $fileInfo.'狀態' = '異常'
                } elseif (-not $validationResult.HashMatch) {
                    $fileInfo.'狀態' = '異常'
                }
                
                $fileInfo.'詳細說明' = $validationResult.Reason
                
                if ($fileInfo.'狀態' -eq '異常') {
                    $mismatchFiles += $fileInfo
                }
            } else {
                $fileInfo.'狀態' = '遺失'
                $fileInfo.'詳細說明' = '備份檔案不存在'
                $mismatchFiles += $fileInfo
            }
            
            $results += $fileInfo
            
            # 每30秒更新一次進度
            $currentTime = Get-Date
            if (($currentTime - $lastProgressUpdate).TotalSeconds -ge 30) {
                $percent = [math]::Round(($processedCount / $total) * 100, 1)
                $elapsedTime = $currentTime - $startTime
                $estimatedTotal = ($elapsedTime.TotalSeconds / $processedCount) * $total
                $remainingTime = $estimatedTotal - $elapsedTime.TotalSeconds
                
                Write-Log "進度: $processedCount / $total ($percent%) - 預估剩餘時間: $('{0:N1}' -f ($remainingTime/60)) 分鐘"
                $lastProgressUpdate = $currentTime
            }
        }
    }
    
    $endTime = Get-Date
    $duration = ($endTime - $startTime).TotalMinutes
    
    # 輸出結果
    Write-Log "`n處理完成:"
    Write-Log "總執行時間: $($duration.ToString('N2')) 分鐘"
    Write-Log "總檔案數: $total"
    Write-Log "正確檔案數: $(($results | Where-Object { $_.狀態 -eq '正確' }).Count)"
    Write-Log "問題檔案數: $($mismatchFiles.Count)"

    # 結束後記錄性能資訊
    $endMemory = Get-CurrentProcessMemory
    $memoryInfo = Get-MemoryUsage
    $cpuUsage = Get-CPUUsage

    Write-Log "`n性能報告:"
    Write-Log "初始處理程序記憶體: $startMemory MB"
    Write-Log "結束處理程序記憶體: $endMemory MB"
    Write-Log "記憶體使用率: $($memoryInfo.MemoryUsagePercent)%"
    Write-Log "系統總記憶體: $($memoryInfo.TotalMemory) MB"
    Write-Log "系統已使用記憶體: $($memoryInfo.UsedMemory) MB"
    Write-Log "平均 CPU 使用率: $cpuUsage%"
   
    # 輸出詳細報告
    $reportPath = Join-Path -Path $LogPath2 -ChildPath "BackupVerificationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSV報告已輸出到: $reportPath"
    
    # 輸出 HTML 報告
    
    $htmlReportPath = Join-Path -Path $LogPath2 -ChildPath "BackupVerificationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    Export-HTMLReport -Results $results -OutputPath $htmlReportPath
    Write-Log "HTML報告已輸出到: $htmlReportPath"
    

    # 如果有問題檔案，輸出單獨的報告
    if ($mismatchFiles.Count -gt 0) {
        $mismatchPath = Join-Path -Path $LogPath2 -ChildPath "ProblemFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $mismatchFiles | Export-Csv -Path $mismatchPath -NoTypeInformation -Encoding UTF8
        Write-Log "問題檔案報告已輸出到: $mismatchPath"
    }

    # 開始前記錄初始狀態
    $startMemory = Get-CurrentProcessMemory
    Write-Log "初始處理程序記憶體: $startMemory MB"


}
catch {
    Write-Log "執行過程中發生錯誤: $($_.Exception.Message)"
    throw
}
pause
