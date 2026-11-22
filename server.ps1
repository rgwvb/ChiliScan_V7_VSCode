# 本機預覽伺服器只綁定回環位址，提供靜態檔案供瀏覽器載入。
$prefix = 'http://127.0.0.1:5500/'
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)

try {
    # 啟動監聽並開啟首頁；工作持續執行直到使用者停止此 PowerShell 工作。
    $listener.Start()
    Write-Host "ChiliScan 預覽伺服器：http://127.0.0.1:5500/"
    Start-Process "http://127.0.0.1:5500/"

    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $requestPath = [Uri]::UnescapeDataString($context.Request.Url.AbsolutePath)
        if ($requestPath -eq '/') { $requestPath = '/index.html' }

        # 正規化請求路徑並確認檔案仍位於專案根目錄，避免讀取專案外的檔案。
        $relativePath = $requestPath.TrimStart('/')
        $root = (Resolve-Path $PSScriptRoot).Path
        $filePath = [System.IO.Path]::GetFullPath((Join-Path $root $relativePath))

        if (-not $filePath.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
            $context.Response.StatusCode = 404
            $context.Response.Close()
            continue
        }

        # 依副檔名設定瀏覽器 MIME 類型，再回傳檔案位元組。
        $extension = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
        $contentTypes = @{
            '.html' = 'text/html; charset=utf-8'
            '.css'  = 'text/css; charset=utf-8'
            '.js'   = 'application/javascript; charset=utf-8'
            '.json' = 'application/json; charset=utf-8'
            '.csv'  = 'text/csv; charset=utf-8'
        }
        $context.Response.ContentType = if ($contentTypes.ContainsKey($extension)) { $contentTypes[$extension] } else { 'application/octet-stream' }
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        $context.Response.StatusCode = 200
        $context.Response.ContentLength64 = $bytes.Length
        $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $context.Response.Close()
    }
} finally {
    # 不論正常停止或發生錯誤，都釋放 HTTP 監聽器資源。
    if ($listener.IsListening) { $listener.Stop() }
    $listener.Close()
}
