$port = 8888
$dir  = Split-Path -Parent $MyInvocation.MyCommand.Path

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-Host ""
Write-Host "  D365 Entity Compare - Server running on http://localhost:$port"
Write-Host "  Serving: $dir"
Write-Host "  Press Ctrl+C to stop."
Write-Host ""

while ($listener.IsListening) {
    $ctx  = $listener.GetContext()
    $req  = $ctx.Request
    $res  = $ctx.Response

    $localPath = $req.Url.LocalPath.TrimStart('/')

    # ── /proxy?url=https://... ──────────────────────────────────────────────
    if ($localPath -eq 'proxy') {
        $rawUrl = $req.QueryString['url']
        $targetUrl = [System.Uri]::UnescapeDataString($rawUrl)
        $rawToken = $req.QueryString['token']
        $bearerToken = if ($rawToken) { [System.Uri]::UnescapeDataString($rawToken) } else { $null }
        Write-Host "  PROXY -> $targetUrl"
        try {
            $wc = New-Object System.Net.WebClient
            $wc.Headers.Add('Accept', 'application/json')
            $wc.Headers.Add('User-Agent', 'D365EntityCompare/1.0')
            if ($bearerToken) {
                $wc.Headers.Add('Authorization', "Bearer $bearerToken")
                Write-Host "  AUTH   Using bearer token (length: $($bearerToken.Length))"
            } else {
                $wc.UseDefaultCredentials = $true
                Write-Host "  AUTH   Using default Windows credentials"
            }
            # Also forward browser cookies if present (helps for forms-based auth)
            $cookieHeader = $req.Headers['Cookie']
            if ($cookieHeader) { $wc.Headers.Add('Cookie', $cookieHeader) }
            $data  = $wc.DownloadString($targetUrl)
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($data)
            $res.StatusCode      = 200
            $res.ContentType     = 'application/json; charset=utf-8'
            $res.Headers.Add('Access-Control-Allow-Origin', '*')
            $res.ContentLength64 = $bytes.Length
            $res.OutputStream.Write($bytes, 0, $bytes.Length)
        } catch {
            $msg = $_.Exception.Message
            $status = 502
            if ($_.Exception -is [System.Net.WebException] -and $_.Exception.Response) {
                $status = [int]$_.Exception.Response.StatusCode
                # Read the actual error body from D365
                try {
                    $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $errBody = $sr.ReadToEnd()
                    $sr.Close()
                    Write-Host "  D365 error body: $errBody"
                } catch {}
            }
            Write-Host "  PROXY ERROR ($status): $msg"
            $body  = [System.Text.Encoding]::UTF8.GetBytes("{`"error`":`"$msg`"}")
            $res.StatusCode      = $status
            $res.ContentType     = 'application/json'
            $res.Headers.Add('Access-Control-Allow-Origin', '*')
            $res.ContentLength64 = $body.Length
            $res.OutputStream.Write($body, 0, $body.Length)
        }
        $res.OutputStream.Close()
        continue
    }

    # ── Static file serving ─────────────────────────────────────────────────
    $path = if ($localPath -eq '' -or $localPath -eq $null) { 'popup.html' } else { $localPath }
    $file = Join-Path $dir $path

    if (Test-Path $file -PathType Leaf) {
        $bytes = [System.IO.File]::ReadAllBytes($file)
        $ext   = [System.IO.Path]::GetExtension($file).ToLower()
        $mime  = switch ($ext) {
            '.html' { 'text/html; charset=utf-8' }
            '.js'   { 'application/javascript' }
            '.css'  { 'text/css' }
            '.json' { 'application/json' }
            default { 'application/octet-stream' }
        }
        $res.StatusCode      = 200
        $res.ContentType     = $mime
        $res.ContentLength64 = $bytes.Length
        $res.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
        $body  = [System.Text.Encoding]::UTF8.GetBytes("404 - Not found: $path")
        $res.StatusCode      = 404
        $res.ContentType     = 'text/plain'
        $res.ContentLength64 = $body.Length
        $res.OutputStream.Write($body, 0, $body.Length)
    }

    $res.OutputStream.Close()
}
