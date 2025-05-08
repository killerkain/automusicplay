<#
.SYNOPSIS
    음악 자동 재생이 가능한 최종 버전 스크립트
.DESCRIPTION
    - Chrome 자동 재생 정책 우회
    - 창 위치 조정 (오른쪽 50%)
    - Windows 7 완벽 호환
#>

# Win32 API 정의
$apiDefinition = @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int nIndex);
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);
}
"@

Add-Type -TypeDefinition $apiDefinition

# 화면 설정
$screenWidth = [Win32]::GetSystemMetrics(0)
$screenHeight = [Win32]::GetSystemMetrics(1)
$halfWidth = [math]::Floor($screenWidth / 2)

# Chrome 실행 함수 (자동 재생 허용)
function Start-ChromeForMusic {
    param ($Url)
    try {
        # 기존 Chrome 프로세스 종료
        Stop-Process -Name "chrome*" -ErrorAction SilentlyContinue
        
        # Chrome 경로 확인
        $chromePath = if (Test-Path "$env:ProgramFiles\Google\Chrome\Application\chrome.exe") {
            "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        } else {
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        }

        if (-not (Test-Path $chromePath)) {
            Write-Output "Chrome이 설치되지 않았습니다."
            exit
        }

        # 자동 재생을 허용하는 Chrome 인수 추가
        $chromeArgs = @(
            "--new-window",
            "--autoplay-policy=no-user-gesture-required",
            "--start-maximized",
            "--window-size=$($halfWidth),$screenHeight",
            "--window-position=$halfWidth,0",
            $Url
        ) -join " "

        # Chrome 실행
        Start-Process -FilePath $chromePath -ArgumentList $chromeArgs

        # 창 위치 조정
        Start-Sleep -Seconds 2
        $hwnd = [Win32]::FindWindow("Chrome_WidgetWin_1", $null)
        if ($hwnd -ne [IntPtr]::Zero) {
            [Win32]::SetWindowPos($hwnd, [IntPtr]::Zero, 
                                $halfWidth, 0, 
                                $halfWidth, $screenHeight, 
                                0x0040)
        }

        Write-Output "Chrome이 음악 재생을 시작했습니다."
    } catch {
        Write-Output "Chrome 실행 오류: $_"
    }
}

# Google 시트 데이터 가져오기
function Get-MusicLinks {
    param ($SheetId, $Range)
    $url = "https://docs.google.com/spreadsheets/d/$SheetId/gviz/tq?tqx=out:csv&range=$Range"
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.Encoding = [System.Text.Encoding]::UTF8
        ($webClient.DownloadString($url)) | ConvertFrom-Csv -Header "Links"
    } catch {
        Write-Output "Google 시트 오류: $_"
        return $null
    }
}

# 메인 실행
$sheetId = "1zjQEDjX6p40xfZ6h0tuxO-YUHqxOS7vss9z3DziKKcA"
$range = "B2:B1000"

while ($true) {
    # 데이터 가져오기
    $links = (Get-MusicLinks -SheetId $sheetId -Range $range | 
             Where-Object { $_.Links -match "^https?://" } | 
             Select-Object -ExpandProperty Links)

    if (-not $links) {
        Write-Output "재생 가능한 링크가 없습니다. 5분 후 재시도..."
        Start-Sleep -Seconds 300
        continue
    }

    # 랜덤 링크 선택
    $musicUrl = $links | Get-Random
    Write-Output "선택된 음악 링크: $musicUrl"

    # Chrome 실행 (자동 재생)
    Start-ChromeForMusic -Url $musicUrl

    # 재생 시간 (5~30분)
    $playTime = Get-Random -Minimum 300 -Maximum 1800
    $endTime = (Get-Date).AddSeconds($playTime)
    Write-Output "재생 중... 종료 시간: $($endTime.ToString('HH:mm:ss'))"

    # 마우스 이동 (활성 상태 유지)
    while ((Get-Date) -lt $endTime) {
        $x = Get-Random -Minimum 100 -Maximum ($screenWidth - 100)
        $y = Get-Random -Minimum 100 -Maximum ($screenHeight - 100)
        [Win32]::SetCursorPos($x, $y)
        Start-Sleep -Seconds (Get-Random -Minimum 15 -Maximum 45)
    }

    # Chrome 종료
    Stop-Process -Name "chrome*" -ErrorAction SilentlyContinue

    # 휴식 시간 (30분~1시간)
    $breakTime = Get-Random -Minimum 1800 -Maximum 3600
    Write-Output "휴식 중... 다음 재생: $((Get-Date).AddSeconds($breakTime).ToString('HH:mm:ss'))"
    Start-Sleep -Seconds $breakTime
}
