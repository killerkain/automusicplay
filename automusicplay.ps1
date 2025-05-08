<#
.SYNOPSIS
    Windows 7 호환 음악 재생 스크립트 (수정버전)
.DESCRIPTION
    - 모든 특수 문자 제거
    - 인코딩 문제 해결
    - PowerShell 2.0 완전 호환
#>

# Windows API for mouse control
$mouseCode = @"
using System;
using System.Runtime.InteropServices;
public class Mouse {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);
}
"@

try {
    Add-Type -TypeDefinition $mouseCode
} catch {
    Write-Host "Mouse control init failed: $_"
}

function Move-RandomMouse {
    $x = Get-Random -Minimum 100 -Maximum 1820
    $y = Get-Random -Minimum 100 -Maximum 980
    try {
        [Mouse]::SetCursorPos($x, $y)
        Write-Host "Mouse moved to ($x, $y)"
    } catch {
        Write-Host "Mouse move error: $_"
    }
}

function Get-GoogleSheetData {
    param (
        [string]$SheetId,
        [string]$Range
    )
    
    $url = "https://docs.google.com/spreadsheets/d/$SheetId/gviz/tq?tqx=out:csv&range=$Range"
    
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.Encoding = [System.Text.Encoding]::UTF8
        $response = $webClient.DownloadString($url)
        return $response | ConvertFrom-Csv -Header "Links"
    } catch {
        Write-Host "Google Sheets error: $_"
        return $null
    }
}

function Start-ChromeWithUrl {
    param ([string]$Url)
    try {
        Stop-Chrome
        $chromePath = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        if (-not (Test-Path $chromePath)) {
            $chromePath = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        }
        Start-Process -FilePath $chromePath -ArgumentList "--new-window $Url --start-maximized"
        Write-Host "Chrome started with URL: $Url"
    } catch {
        Write-Host "Chrome start error: $_"
    }
}

function Stop-Chrome {
    try {
        Get-Process -Name "chrome*" -ErrorAction SilentlyContinue | Stop-Process -Force
    } catch {
        Write-Host "Chrome stop error: $_"
    }
}

# Main execution
$sheetId = "1zjQEDjX6p40xfZ6h0tuxO-YUHqxOS7vss9z3DziKKcA"
$range = "B2:B1000"

while ($true) {
    $sheetData = Get-GoogleSheetData -SheetId $sheetId -Range $range
    $links = $sheetData | Where-Object { $_.Links -match "^https?://" } | Select-Object -ExpandProperty Links
    
    if (-not $links -or $links.Count -eq 0) {
        Write-Host "No valid links found. Retrying in 5 minutes..."
        Start-Sleep -Seconds 300
        continue
    }

    $randomLink = $links | Get-Random
    Start-ChromeWithUrl -Url $randomLink
    
    $playTime = Get-Random -Minimum 300 -Maximum 1800
    $endTime = (Get-Date).AddSeconds($playTime)
    
    Write-Host "Playback started | End time: $($endTime.ToString('HH:mm:ss'))"
    
    while ((Get-Date) -lt $endTime) {
        Move-RandomMouse
        Start-Sleep -Seconds (Get-Random -Minimum 15 -Maximum 45)
    }
    
    Stop-Chrome
    
    $breakTime = Get-Random -Minimum 1800 -Maximum 3600
    Write-Host "Break started | Resume time: $((Get-Date).AddSeconds($breakTime).ToString('HH:mm:ss'))"
    Start-Sleep -Seconds $breakTime
}
