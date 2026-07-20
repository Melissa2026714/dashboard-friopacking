@echo off
title Dashboard Compras FRIOPACKING
color 1F
echo.
echo  ========================================
echo   FRIOPACKING S.A. - Dashboard de Compras
echo  ========================================
echo.
echo  Iniciando servidor local...
echo  NO cierres esta ventana mientras uses el dashboard.
echo.

set "SCRIPT_DIR=%~dp0"
set "BASE_DIR=%SCRIPT_DIR%.."
set "PORT=8080"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$port=%PORT%; $base=(Resolve-Path '%BASE_DIR%').Path;" ^
  "$listener = New-Object System.Net.HttpListener;" ^
  "$listener.Prefixes.Add('http://localhost:' + $port + '/');" ^
  "try { $listener.Start() } catch { Write-Host ' ERROR: Puerto ocupado. Cierra el navegador y reintenta.' -ForegroundColor Red; Start-Sleep 4; exit };" ^
  "Write-Host ' Servidor activo en http://localhost:' $port -ForegroundColor Green;" ^
  "Write-Host ' Abriendo navegador...' -ForegroundColor Cyan;" ^
  "Start-Process ('http://localhost:' + $port + '/plataforma.html');" ^
  "$mime = @{'.html'='text/html; charset=utf-8';'.htm'='text/html; charset=utf-8';'.js'='application/javascript; charset=utf-8';'.json'='application/json; charset=utf-8';'.css'='text/css; charset=utf-8';'.png'='image/png';'.jpg'='image/jpeg';'.jpeg'='image/jpeg';'.gif'='image/gif';'.svg'='image/svg+xml';'.ico'='image/x-icon';'.xlsx'='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';'.pdf'='application/pdf' };" ^
  "while ($listener.IsListening) {" ^
  "  try {" ^
  "    $ctx = $listener.GetContext();" ^
  "    $reqPath = [System.Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath);" ^
  "    if ($reqPath -eq '/') { $reqPath = '/index.html' }" ^
  "    $rel = $reqPath.TrimStart('/').Replace('/', [System.IO.Path]::DirectorySeparatorChar);" ^
  "    $file = Join-Path $base $rel;" ^
  "    $fullFile = [System.IO.Path]::GetFullPath($file);" ^
  "    if (-not $fullFile.StartsWith($base) -or -not (Test-Path $fullFile -PathType Leaf)) { $ctx.Response.StatusCode = 404; $ctx.Response.Close(); continue; }" ^
  "    $ext = [System.IO.Path]::GetExtension($fullFile).ToLower();" ^
  "    $ct = $mime[$ext]; if (-not $ct) { $ct = 'application/octet-stream' };" ^
  "    $content = [System.IO.File]::ReadAllBytes($fullFile);" ^
  "    $ctx.Response.ContentType = $ct;" ^
  "    $ctx.Response.ContentLength64 = $content.Length;" ^
  "    $ctx.Response.Headers.Add('Cache-Control','no-cache');" ^
  "    $ctx.Response.OutputStream.Write($content, 0, $content.Length);" ^
  "    $ctx.Response.Close();" ^
  "  } catch {}" ^
  "}"
