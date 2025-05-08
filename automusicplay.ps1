<#
.SYNOPSIS
    Windows 7 완전 호환 음악 재생 스크립트
.DESCRIPTION
    - 모든 유니코드 문자 제거
    - UTF-8 BOM 인코딩 적용
    - PowerShell 2.0 완벽 지원
#>

# 마우스 제어 API
$mouseCode = @'
using System;
using System.Runtime.InteropServices;
public class Mouse {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);
}
'@

try { Add-Type -TypeDefinition $mouseCode } catch { Write-Output "Mouse API load failed" }

function Move-MouseRandom {
    $x = Get-Random -Minimum 100 -Maximum 1820
    $y = Get-Random -Minimum 100 -Maximum 980
    try {
        [Mouse]::SetCursorPos($x, $y)
        Write-Output "Mouse moved to ($x, $y)"
    } catch { Write-Output "Mouse move error" }
}

function Get-SheetData {
    param ($SheetId, $Range)
    
    $url = "https://docs.google.com/spreadsheets/d/$SheetId/gviz/tq?tqx=out:csv&range=$Range"
    
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Encoding = [System.Text.Encoding]::UTF8
        $data = $wc.DownloadString($url) | ConvertFrom-Csv -Header "Links"
        return $data
    } catch {
        Write-Output "Google Sheets error"
        return $null
    }
}

function Start-Chrome {
    param ($Url)
    try {
        Stop-Chrome
        $chrome = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        if (!(Test-Path $chrome)) { $chrome = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe" }
        Start-Process $chrome "--new-window $Url --start-maximized"
        Write-Output "Chrome started: $Url"
    } catch { Write-Output "Chrome start error" }
}

function Stop-Chrome {
    try { Stop-Process -Name "chrome*" -Force -ErrorAction SilentlyContinue }
    catch { Write-Output "Chrome stop error" }
}

# 메인 실행
$sheetId = "1zjQEDjX6p40xfZ6h0tuxO-YUHqxOS7vss9z3DziKKcA"
$range = "B2:B1000"

while ($true) {
    $data = Get-SheetData -SheetId $sheetId -Range $range
    $links = $data | Where { $_.Links -match "^https?://" } | Select -ExpandProperty Links
    
    if (!$links -or $links.Count -eq 0) {
        Write-Output "No links found. Retry in 5 minutes..."
        Start-Sleep -Seconds 300
        continue
    }

    $url = $links | Get-Random
    Start-Chrome -Url $url
    
    $playSec = Get-Random -Minimum 300 -Maximum 1800
    $endTime = (Get-Date).AddSeconds($playSec)
    
    Write-Output "Playing until $($endTime.ToString('HH:mm:ss'))"
    
    while ((Get-Date) -lt $endTime) {
        Move-MouseRandom
        Start-Sleep -Seconds (Get-Random -Minimum 15 -Maximum 45)
    }
    
    Stop-Chrome
    
    $breakSec = Get-Random -Minimum 1800 -Maximum 3600
    Write-Output "Break until $((Get-Date).AddSeconds($breakSec).ToString('HH:mm:ss'))"
    Start-Sleep -Seconds $breakSec
}
