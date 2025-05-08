<#
.SYNOPSIS
    최종 버전: Chrome 음악 재생 + 랜덤 마우스 이동
.DESCRIPTION
    - Chrome 브라우저 사용
    - 음악 재생: 5~30분 (300~1800초)
    - 휴식 시간: 30분~1시간 (1800~3600초)
    - 1920x1080 화면 내 랜덤 마우스 이동
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
    $x = Get-Random -Minimum 100 -Maximum 1820  # 작업표시줄 영역 제외
    $y = Get-Random -Minimum 100 -Maximum 980   # 시작 메뉴 영역 제외
    [void][Mouse]::SetCursorPos($x, $y)
    Write-Host "마우스 이동: ($x, $y) - 남은 시간: $(([DateTime]::Now - $endTime).ToString('mm\:ss')"
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
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Google 시트 오류: $_"
        return $null
    }
}

function Start-ChromeWithUrl {
    param ([string]$Url)
    try {
        Stop-Chrome
        # Chrome 단일 창 모드 실행 (기존 창 재사용 방지)
        Start-Process "chrome.exe" "--new-window $Url --start-maximized"
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Chrome 재생 시작: $Url"
    } catch {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Chrome 실행 오류: $_"
        exit
    }
}

function Stop-Chrome {
    try {
        Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Chrome 종료 오류: $_"
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
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 재생 가능한 링크 없음. 5분 후 재시도..."
        Start-Sleep -Seconds 300
        continue
    }

    # 음악 재생
    $randomLink = $links | Get-Random
    Start-ChromeWithUrl -Url $randomLink
    
    # 재생 시간 설정 (5~30분)
    $playTime = Get-Random -Minimum 300 -Maximum 1800
    $endTime = [DateTime]::Now.AddSeconds($playTime)
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 재생 시작 | 종료 예정: $($endTime.ToString('HH:mm:ss'))"
    
    # 재생 중 마우스 이동 (15~45초 간격)
    while ([DateTime]::Now -lt $endTime) {
        Move-RandomMouse
        $sleepTime = Get-Random -Minimum 15 -Maximum 45
        Start-Sleep -Seconds $sleepTime
    }
    
    Stop-Chrome
    
    # 휴식 시간 (30분~1시간)
    $breakTime = Get-Random -Minimum 1800 -Maximum 3600
    $restEndTime = [DateTime]::Now.AddSeconds($breakTime)
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 휴식 시작 | 재개 예정: $($restEndTime.ToString('HH:mm:ss'))`n"
    Start-Sleep -Seconds $breakTime
}
