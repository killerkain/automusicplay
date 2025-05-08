<#
.SYNOPSIS
    음악 재생 + 랜덤 마우스 이동 스크립트
.DESCRIPTION
    - Firefox 단일 창에서 음악 재생
    - 1920x1080 화면 내에서 랜덤 마우스 이동
    - 테스트용 1~2분 재생, 2~5분 휴식
#>

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Mouse {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);
}
"@

function Move-RandomMouse {
    $x = Get-Random -Minimum 0 -Maximum 1920
    $y = Get-Random -Minimum 0 -Maximum 1080
    [void][Mouse]::SetCursorPos($x, $y)
    Write-Host "마우스 이동: ($x, $y)"
}

function Get-GoogleSheetData {
    param (
        [string]$SheetId,
        [string]$Range
    )
    $url = "https://docs.google.com/spreadsheets/d/$SheetId/gviz/tq?tqx=out:csv&range=$Range"
    try {
        $response = Invoke-RestMethod -Uri $url -Method Get
        $data = $response | ConvertFrom-Csv -Header "Links"
        return $data
    } catch {
        Write-Host "Google 시트 오류: $_"
        return $null
    }
}

function Start-FirefoxWithUrl {
    param ([string]$Url)
    try {
        Stop-Firefox
        Start-Process "firefox.exe" "-new-window $Url"
        Write-Host "Firefox 시작: $Url"
    } catch {
        Write-Host "Firefox 실행 오류: $_"
        exit
    }
}

function Stop-Firefox {
    try {
        Stop-Process -Name "firefox" -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "Firefox 종료 오류: $_"
    }
}

# 메인 실행
$sheetId = "1zjQEDjX6p40xfZ6h0tuxO-YUHqxOS7vss9z3DziKKcA"
$range = "B2:B1000"

while ($true) {
    # 데이터 가져오기
    $sheetData = Get-GoogleSheetData -SheetId $sheetId -Range $range
    $links = $sheetData | Where-Object { $_.Links -match "^https?://" } | Select-Object -ExpandProperty Links
    
    if ($links.Count -eq 0) {
        Write-Host "링크 없음. 1분 후 재시도..."
        Start-Sleep -Seconds 60
        continue
    }

    # 음악 재생
    $randomLink = $links | Get-Random
    Start-FirefoxWithUrl -Url $randomLink
    
    # 재생 시간 설정 (1~2분)
    $playTime = Get-Random -Minimum 60 -Maximum 120
    $endTime = [DateTime]::Now.AddSeconds($playTime)
    
    Write-Host "재생 시간: $playTime 초 ($($endTime.ToString('HH:mm:ss'))"
    
    # 재생 중 마우스 이동
    while ([DateTime]::Now -lt $endTime) {
        Move-RandomMouse
        $sleepTime = Get-Random -Minimum 10 -Maximum 30
        Start-Sleep -Seconds $sleepTime
    }
    
    Stop-Firefox
    
    # 휴식 시간 (2~5분)
    $breakTime = Get-Random -Minimum 120 -Maximum 300
    Write-Host "휴식 시간: $breakTime 초`n"
    Start-Sleep -Seconds $breakTime
}