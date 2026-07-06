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
  "$port=%PORT%; $base='%BASE_DIR%';" ^
  "$listener = New-Object System.Net.HttpListener;" ^
  "$listener.Prefixes.Add('http://localhost:' + $port + '/');" ^
  "try { $listener.Start() } catch { Write-Host ' ERROR: Puerto ocupado. Cierra el navegador y reintenta.' -ForegroundColor Red; Start-Sleep 4; exit };" ^
  "Write-Host ' Servidor activo en http://localhost:' $port -ForegroundColor Green;" ^
  "Write-Host ' Abriendo navegador...' -ForegroundColor Cyan;" ^
  "Start-Process ('http://localhost:' + $port);" ^
  "while ($listener.IsListening) {" ^
  "  try {" ^
  "    $ctx = $listener.GetContext();" ^
  "    $path = $ctx.Request.Url.AbsolutePath;" ^
  "    if ($path -eq '/' -or $path -eq '/index.html') {" ^
  "      $file = $base + '\index.html';" ^
  "      $ct = 'text/html; charset=utf-8';" ^
  "    } elseif ($path -match '\\.js$') {" ^
  "      $name = [System.IO.Path]::GetFileName($path);" ^
  "      $file = $base + '\' + $name;" ^
  "      $ct = 'application/javascript; charset=utf-8';" ^
  "    } else {" ^
  "      $ctx.Response.StatusCode = 404;" ^
  "      $ctx.Response.Close();" ^
  "      continue;" ^
  "    }" ^
  "    if (-not (Test-Path $file)) { $ctx.Response.StatusCode=404; $ctx.Response.Close(); continue; }" ^
  "    $content = [System.IO.File]::ReadAllBytes($file);" ^
  "    $ctx.Response.ContentType = $ct;" ^
  "    $ctx.Response.ContentLength64 = $content.Length;" ^
  "    $ctx.Response.Headers.Add('Cache-Control','no-cache');" ^
  "    $ctx.Response.OutputStream.Write($content, 0, $content.Length);" ^
  "    $ctx.Response.Close();" ^
  "  } catch {}" ^
  "}"
