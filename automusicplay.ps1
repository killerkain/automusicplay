<#
.SYNOPSIS
    Chrome을 화면 오른쪽 절반에 열어주는 음악 재생 스크립트
.DESCRIPTION
    - Chrome 창이 화면 오른쪽 50% 영역에 정확히 배치
    - 기존 모든 기능 유지 (랜덤 마우스 이동, 재생/휴식 시간 등)
#>

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Window {
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
}
"@

# 화면 해상도 얻기
$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
$halfWidth = $screenWidth / 2

function Start-ChromeHalfScreen {
    param ($Url)
    try {
        Stop-Chrome
        
        # Chrome 실행 (최소화 상태로 시작)
        $chromePath = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        if (-not (Test-Path $chromePath)) {
            $chromePath = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        }
        $process = Start-Process $chromePath "--new-window $Url" -PassThru
        
        # 3초 대기 후 창 조정
        Start-Sleep -Seconds 3
        
        # Chrome 창 핸들 찾기
        $hwnd = [Window]::FindWindow("Chrome_WidgetWin_1", $null)
        if ($hwnd -ne [IntPtr]::Zero) {
            # 오른쪽 절반으로 창 이동 (X: 화면너비/2, Y: 0, Width: 화면너비/2, Height: 전체높이)
            [Window]::SetWindowPos($hwnd, [IntPtr]::Zero, 
                                 $halfWidth, 0, 
                                 $halfWidth, $screenHeight, 
                                 0x0040)  # 0x0040 = SWP_SHOWWINDOW
            
            Write-Output "Chrome을 오른쪽 절반에 배치했습니다."
        } else {
            Write-Output "Chrome 창 핸들을 찾을 수 없습니다."
        }
    } catch {
        Write-Output "Chrome 실행 오류: $_"
    }
}

# 기존 함수들 유지 (Get-SheetData, Move-MouseRandom, Stop-Chrome 등은 동일)

# 메인 실행부
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
    Start-ChromeHalfScreen -Url $url  # ← 수정된 함수 호출
    
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
