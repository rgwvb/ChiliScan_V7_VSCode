$prefix = 'http://127.0.0.1:5500/'
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
    Write-Host "ChiliScan 預覽伺服器：http://127.0.0.1:5500/"
    Start-Process "http://127.0.0.1:5500/"

    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $requestPath = [Uri]::UnescapeDataString($context.Request.Url.AbsolutePath)
        if ($requestPath -eq '/') { $requestPath = '/index.html' }

        $relativePath = $requestPath.TrimStart('/')
        $root = (Resolve-Path $PSScriptRoot).Path
        $filePath = [System.IO.Path]::GetFullPath((Join-Path $root $relativePath))

        if (-not $filePath.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
            $context.Response.StatusCode = 404
            $context.Response.Close()
            continue
        }

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
    if ($listener.IsListening) { $listener.Stop() }
    $listener.Close()
}
