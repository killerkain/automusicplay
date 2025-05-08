<#
.SYNOPSIS
    Chrome 음악 자동 재생 스크립트 (마우스 클릭 없음)
.DESCRIPTION
    - Chrome으로 음악 자동 재생 (마우스 클릭 없이)
    - 창을 오른쪽 절반에 고정
    - 랜덤 마우스 이동으로 시스템 절전 모드 방지
    - Windows 7 완벽 호환
#>

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int nIndex);
}
"@

# 화면 해상도 설정
$screenWidth = [Win32]::GetSystemMetrics(0)
$screenHeight = [Win32]::GetSystemMetrics(1)
$halfWidth = [math]::Floor($screenWidth / 2)

# Chrome 실행 (자동 재생 강제 활성화)
function Start-ChromeMusic {
    param($Url)
    try {
        # 기존 Chrome 종료
        Stop-Process -Name "chrome*" -ErrorAction SilentlyContinue

        # Chrome 경로 확인
        $chromePath = if (Test-Path "$env:ProgramFiles\Google\Chrome\Application\chrome.exe") {
            "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        } else {
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        }

        if (-not (Test-Path $chromePath)) {
            Write-Host "Chrome이 설치되지 않았습니다." -ForegroundColor Red
            exit
        }

        # Chrome 실행 (자동 재생 정책 우회 + 시크릿 모드)
        $process = Start-Process $chromePath @(
            "--new-window",
            "--autoplay-policy=no-user-gesture-required",
            "--disable-features=PreloadMediaEngagementData,AutoplayIgnoreWebAudio",
            "--window-size=$halfWidth,$screenHeight",
            "--window-position=$halfWidth,0",
            $Url
        ) -PassThru

        # 창 핸들 찾아 위치 고정
        Start-Sleep -Seconds 3
        $hwnd = [Win32]::FindWindow("Chrome_WidgetWin_1", $null)
        if ($hwnd -ne [IntPtr]::Zero) {
            [Win32]::SetWindowPos($hwnd, [IntPtr]::Zero, $halfWidth, 0, $halfWidth, $screenHeight, 0x0040)
        }

        Write-Host "Chrome에서 음악 재생을 시작했습니다." -ForegroundColor Green
        return $process
    } catch {
        Write-Host "오류: $_" -ForegroundColor Red
        return $null
    }
}

# Google 시트에서 링크 가져오기
function Get-MusicLinks {
    param($SheetId, $Range)
    $url = "https://docs.google.com/spreadsheets/d/$SheetId/gviz/tq?tqx=out:csv&range=$Range"
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.Encoding = [System.Text.Encoding]::UTF8
        ($webClient.DownloadString($url)) | ConvertFrom-Csv -Header "Links"
    } catch {
        Write-Host "Google 시트 오류: $_" -ForegroundColor Yellow
        return $null
    }
}

# 메인 실행
$sheetId = "1zjQEDjX6p40xfZ6h0tuxO-YUHqxOS7vss9z3DziKKcA"
$range = "B2:B1000"

while ($true) {
    # 링크 가져오기
    $links = (Get-MusicLinks -SheetId $sheetId -Range $range | 
             Where-Object { $_.Links -match "^https?://" } | 
             Select-Object -ExpandProperty Links)

    if (-not $links) {
        Write-Host "재생 가능한 링크가 없습니다. 5분 후 재시도..." -ForegroundColor Yellow
        Start-Sleep -Seconds 300
        continue
    }

    # 랜덤 링크 선택
    $musicUrl = $links | Get-Random
    Write-Host "선택된 음악 링크: $musicUrl" -ForegroundColor Cyan

    # Chrome 실행 및 음악 재생
    $chromeProcess = Start-ChromeMusic -Url $musicUrl

    if (-not $chromeProcess) {
        Write-Host "재생 실패. 5분 후 재시도..." -ForegroundColor Red
        Start-Sleep -Seconds 300
        continue
    }

    # 재생 시간 (5~30분)
    $playTime = Get-Random -Minimum 300 -Maximum 1800
    $endTime = (Get-Date).AddSeconds($playTime)
    Write-Host "재생 중... 종료 시간: $($endTime.ToString('HH:mm:ss'))" -ForegroundColor Green

    # 마우스 이동 (15~45초 간격) - 클릭 없음!
    while ((Get-Date) -lt $endTime) {
        $x = Get-Random -Minimum ($halfWidth + 50) -Maximum ($screenWidth - 50)
        $y = Get-Random -Minimum 50 -Maximum ($screenHeight - 50)
        [Win32]::SetCursorPos($x, $y)
        Start-Sleep -Seconds (Get-Random -Minimum 15 -Maximum 45)
    }

    # Chrome 종료
    Stop-Process -Id $chromeProcess.Id -ErrorAction SilentlyContinue

    # 휴식 시간 (30분~1시간)
    $breakTime = Get-Random -Minimum 1800 -Maximum 3600
    Write-Host "휴식 중... 다음 재생: $((Get-Date).AddSeconds($breakTime).ToString('HH:mm:ss'))`n" -ForegroundColor Magenta
    Start-Sleep -Seconds $breakTime
}
