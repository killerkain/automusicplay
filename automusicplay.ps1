<#
.SYNOPSIS
    Windows 7 완벽 호환 음악 재생 스크립트
.DESCRIPTION
    - Chrome을 오른쪽 절반 화면에 표시
    - .NET Framework 2.0 호환
    - PowerShell 2.0 기본 기능만 사용
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

try {
    Add-Type -TypeDefinition $apiDefinition
} catch {
    Write-Output "Win32 API 로드 실패: $_"
    exit
}

# 화면 해상도 얻기
$SM_CXSCREEN = 0
$SM_CYSCREEN = 1
$screenWidth = [Win32]::GetSystemMetrics($SM_CXSCREEN)
$screenHeight = [Win32]::GetSystemMetrics($SM_CYSCREEN)
$halfWidth = [math]::Floor($screenWidth / 2)

# Chrome 실행 함수
function Start-ChromeHalfScreen {
    param ($Url)
    try {
        # 기존 Chrome 프로세스 종료
        Stop-Process -Name "chrome*" -ErrorAction SilentlyContinue
        
        # Chrome 경로 확인
        $chromePath = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        if (-not (Test-Path $chromePath)) {
            $chromePath = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        }
        
        if (-not (Test-Path $chromePath)) {
            Write-Output "Chrome을 찾을 수 없습니다. 설치 후 다시 시도하세요."
            exit
        }
        
        # Chrome 실행
        $process = Start-Process -FilePath $chromePath -ArgumentList "--new-window $Url" -PassThru
        
        # 창 핸들 찾기를 위한 대기
        Start-Sleep -Seconds 3
        
        # 창 핸들 찾기
        $hwnd = [Win32]::FindWindow("Chrome_WidgetWin_1", $null)
        if ($hwnd -ne [IntPtr]::Zero) {
            # 창 크기/위치 조정 (오른쪽 절반)
            [Win32]::SetWindowPos($hwnd, [IntPtr]::Zero, 
                                $halfWidth, 0, 
                                $halfWidth, $screenHeight, 
                                0x0040)  # SWP_SHOWWINDOW
            Write-Output "Chrome을 오른쪽 절반에 배치했습니다."
        } else {
            Write-Output "Chrome 창을 찾을 수 없습니다."
        }
    } catch {
        Write-Output "Chrome 실행 오류: $_"
    }
}

# Google 시트 데이터 가져오기
function Get-SheetData {
    param ($SheetId, $Range)
    
    $url = "https://docs.google.com/spreadsheets/d/$SheetId/gviz/tq?tqx=out:csv&range=$Range"
    
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.Encoding = [System.Text.Encoding]::UTF8
        $response = $webClient.DownloadString($url)
        return $response | ConvertFrom-Csv -Header "Links"
    } catch {
        Write-Output "Google 시트 접근 오류: $_"
        return $null
    }
}

# 랜덤 마우스 이동
function Move-RandomMouse {
    $x = Get-Random -Minimum 100 -Maximum ($screenWidth - 100)
    $y = Get-Random -Minimum 100 -Maximum ($screenHeight - 100)
    [Win32]::SetCursorPos($x, $y)
    Write-Output "마우스 이동: ($x, $y)"
}

# 메인 실행
$sheetId = "1zjQEDjX6p40xfZ6h0tuxO-YUHqxOS7vss9z3DziKKcA"
$range = "B2:B1000"

while ($true) {
    # 데이터 가져오기
    $data = Get-SheetData -SheetId $sheetId -Range $range
    if (-not $data) {
        Write-Output "데이터를 가져오지 못했습니다. 5분 후 재시도..."
        Start-Sleep -Seconds 300
        continue
    }
    
    $links = $data | Where-Object { $_.Links -match "^https?://" } | Select-Object -ExpandProperty Links
    if (-not $links -or $links.Count -eq 0) {
        Write-Output "재생 가능한 링크가 없습니다. 5분 후 재시도..."
        Start-Sleep -Seconds 300
        continue
    }

    # 랜덤 링크 선택
    $url = $links | Get-Random
    Write-Output "재생할 링크 선택됨: $url"
    Start-ChromeHalfScreen -Url $url
    
    # 재생 시간 (5~30분)
    $playTime = Get-Random -Minimum 300 -Maximum 1800
    $endTime = (Get-Date).AddSeconds($playTime)
    Write-Output "재생 중... 종료 예정: $($endTime.ToString('HH:mm:ss'))"
    
    # 마우스 이동 (15~45초 간격)
    while ((Get-Date) -lt $endTime) {
        Move-RandomMouse
        $sleepTime = Get-Random -Minimum 15 -Maximum 45
        Start-Sleep -Seconds $sleepTime
    }
    
    # Chrome 종료
    Stop-Process -Name "chrome*" -ErrorAction SilentlyContinue
    
    # 휴식 시간 (30분~1시간)
    $breakTime = Get-Random -Minimum 1800 -Maximum 3600
    Write-Output "휴식 중... 재개 예정: $((Get-Date).AddSeconds($breakTime).ToString('HH:mm:ss'))"
    Start-Sleep -Seconds $breakTime
}
