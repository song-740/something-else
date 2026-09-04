# something-else

# 1. Script Math-Type

Syntax:

```sh
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Users\admin\Desktop\auto\[1.1] Fix.ps1" -DocumentPath "C:\Users\admin\Desktop\auto\tmp.docx" -OutputPath "C:\Users\admin\Desktop\auto\tmp-mathtype.docx"
```

```sh
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$DocumentPath,

    [Parameter(Position = 1)]
    [string]$OutputPath,

    [ValidateRange(100, 10000)]
    [int]$WaitMilliseconds = 900
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WordWindowNative {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

$wdDoNotSaveChanges = 0
$wdAlertsNone = 0
$wdWindowStateNormal = 0
$swRestore = 9

$word = $null
$document = $null
$window = $null
$shell = $null
$storyRanges = $null
$documents = $null
$addIns = $null
$mathTypeAddIn = $null
$equations = New-Object System.Collections.ArrayList
$createdOutput = $false
$savedOutput = $false
$successCount = 0
$failureCount = 0
$fatalError = $null
$wordProcessId = 0

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Host ('[{0}] [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message)
}

function Get-AbsolutePath {
    param([Parameter(Mandatory = $true)][string]$Path)
    return [System.IO.Path]::GetFullPath($Path)
}

function Release-ComObject {
    param($ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) } catch {}
    }
}

function Get-Preview {
    param([string]$Text)
    $preview = $Text.Replace("`r", ' ').Replace("`n", ' ')
    if ($preview.Length -gt 90) { return $preview.Substring(0, 87) + '...' }
    return $preview
}

# function Activate-WordWindow {
#     $document.Activate()
#     $window.Activate()
#     $word.WindowState = $wdWindowStateNormal
#     $word.Visible = $true
#     $handle = [IntPtr]([int64]$window.Hwnd)
#     [void][WordWindowNative]::ShowWindowAsync($handle, $swRestore)
#     if (-not [WordWindowNative]::SetForegroundWindow($handle)) {
#         [void]$shell.AppActivate([string]$window.Caption)
#     }
#     Start-Sleep -Milliseconds 120
# }

function Activate-WordWindow {
    [void]$document.Activate()
    [void]$window.Activate()

    $word.WindowState = $wdWindowStateNormal
    $word.Visible = $true

    $handle = [IntPtr]([int64]$window.Hwnd)

    [void][WordWindowNative]::ShowWindowAsync(
        $handle,
        $swRestore
    )

    if (-not [WordWindowNative]::SetForegroundWindow($handle)) {
        [void]$shell.AppActivate([string]$window.Caption)
    }

    Start-Sleep -Milliseconds 120
}

function Test-RangeConverted {
    param($Range, [string]$OriginalText)
    $currentText = [string]$Range.Text
    return (-not $currentText.Contains($OriginalText)) -and
           (-not [regex]::IsMatch($currentText, '^\$[^\r\n\a]*\$$'))
}

# MathType Toggle TeX may ignore expressions that contain multiple
# characters but no explicit operator, such as:
#
#   $AB$
#   $22$
#   $12x$
#
# Convert those expressions into more explicit TeX forms only after
# the normal Toggle TeX attempts have failed.

function Get-DollarTeXFallback {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    # Accept exactly one non-empty inline expression using single dollars.
    $outerMatch = [regex]::Match(
        $Text,
        '^\$(?<Inner>[^\$\r\n\a]+)\$$'
    )

    if (-not $outerMatch.Success) {
        return $null
    }

    $innerText = $outerMatch.Groups['Inner'].Value

    # Multiple Latin letters:
    #
    #   $AB$   -> $\mathit{AB}$
    #   $AEHF$ -> $\mathit{AEHF}$
    if ($innerText -match '^[A-Za-z]{2,}$') {
        return [pscustomobject]@{
            Category     = 'PlainMultiLetter'
            OriginalText = $Text
            InnerText    = $innerText
            FallbackText = '$\mathit{' + $innerText + '}$'
        }
    }

    # Multiple digits:
    #
    #   $22$   -> $\mathrm{22}$
    #   $2026$ -> $\mathrm{2026}$
    #
    # Do not apply this to $2$, because a single character normally works.
    if ($innerText -match '^\d{2,}$') {
        return [pscustomobject]@{
            Category     = 'PlainMultiDigit'
            OriginalText = $Text
            InnerText    = $innerText
            FallbackText = '$\mathrm{' + $innerText + '}$'
        }
    }

    # Numeric coefficient followed by Latin variables:
    #
    #   $12x$  -> $12\mathit{x}$
    #   $2xy$  -> $2\mathit{xy}$
    #   $12ABC$ -> $12\mathit{ABC}$
    #
    # Keep the number in normal mathematical number style and explicitly
    # mark the letter portion as mathematical italic.
    $coefficientVariableMatch = [regex]::Match(
        $innerText,
        '^(?<Coefficient>\d+)(?<Variables>[A-Za-z]+)$'
    )

    if ($coefficientVariableMatch.Success) {
        $coefficient = $coefficientVariableMatch.Groups['Coefficient'].Value
        $variables = $coefficientVariableMatch.Groups['Variables'].Value

        return [pscustomobject]@{
            Category     = 'CoefficientVariable'
            OriginalText = $Text
            InnerText    = $innerText
            FallbackText = (
                '$' +
                $coefficient +
                '\mathit{' +
                $variables +
                '}$'
            )
        }
    }

    return $null
}

function Invoke-DollarTeXFallback {
    param(
        $Range,

        [Parameter(Mandatory = $true)]
        [string]$OriginalText,

        [Parameter(Mandatory = $true)]
        $Fallback,

        [Parameter(Mandatory = $true)]
        [int]$WaitMilliseconds
    )

    $fallbackText = [string]$Fallback.FallbackText
    $category = [string]$Fallback.Category

    if ([string]::IsNullOrWhiteSpace($fallbackText)) {
        return [pscustomobject]@{
            Converted = $false
            Attempted = $false
            Category  = $category
            Message   = 'Fallback text is empty.'
        }
    }

    try {
        # Prevent all COM method results from leaking into the PowerShell
        # function output. The function must return exactly one object.
        [void]($Range.Text = $fallbackText)

        Activate-WordWindow

        [void]$Range.Select()

        Activate-WordWindow

        $ribbonControl = $null

        [void]$word.Run(
            'MTCommand_OnTexToggle',
            [ref]$ribbonControl
        )

        Start-Sleep -Milliseconds $WaitMilliseconds

        $converted = Test-RangeConverted `
            -Range $Range `
            -OriginalText $fallbackText

        if ($converted) {
            return [pscustomobject]@{
                Converted = $true
                Attempted = $true
                Category  = $category
                Message   = (
                    "Converted by {0} fallback: {1}" -f
                    $category,
                    $fallbackText
                )
            }
        }

        # Fallback did not convert. Restore the exact original TeX.
        [void]($Range.Text = $OriginalText)

        return [pscustomobject]@{
            Converted = $false
            Attempted = $true
            Category  = $category
            Message   = (
                "{0} fallback did not replace the text: {1}" -f
                $category,
                $fallbackText
            )
        }
    }
    catch {
        $fallbackError = [string]$_.Exception.Message

        try {
            [void]($Range.Text = $OriginalText)
        }
        catch {
            # Keep the original fallback exception as the reported error.
        }

        return [pscustomobject]@{
            Converted = $false
            Attempted = $true
            Category  = $category
            Message   = (
                "{0} fallback failed: {1}" -f
                $category,
                $fallbackError
            )
        }
    }
}


try {
    $DocumentPath = Get-AbsolutePath $DocumentPath
    if (-not (Test-Path -LiteralPath $DocumentPath -PathType Leaf)) {
        throw "Input file not found: $DocumentPath"
    }
    if (-not ([System.IO.Path]::GetExtension($DocumentPath) -ieq '.docx')) {
        throw "Input file must have the .docx extension: $DocumentPath"
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $directory = [System.IO.Path]::GetDirectoryName($DocumentPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($DocumentPath)
        $OutputPath = Join-Path $directory ($baseName + '_mathtype.docx')
    }
    $OutputPath = Get-AbsolutePath $OutputPath
    if (-not ([System.IO.Path]::GetExtension($OutputPath) -ieq '.docx')) {
        throw "Output file must have the .docx extension: $OutputPath"
    }
    if ($DocumentPath.Equals($OutputPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'DocumentPath and OutputPath must differ; the input file will not be overwritten.'
    }

    $outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        [void][System.IO.Directory]::CreateDirectory($outputDirectory)
    }

    Copy-Item -LiteralPath $DocumentPath -Destination $OutputPath -Force
    $createdOutput = $true
    Write-Log "Working copy created: $OutputPath"

    try {
        $word = New-Object -ComObject Word.Application
    }
    catch {
        throw "Cannot start Microsoft Word COM Automation. Verify that Word is installed and working. Details: $($_.Exception.Message)"
    }

    $word.Visible = $true
    $word.DisplayAlerts = $wdAlertsNone
    $word.AutomationSecurity = 1
    $word.WindowState = $wdWindowStateNormal
    $shell = New-Object -ComObject WScript.Shell

    $mathTypeTemplateCandidates = @(
        'C:\Program Files (x86)\MathType\Office Support\64\MathType Commands 2016.dotm',
        'C:\Program Files\MathType\Office Support\64\MathType Commands 2016.dotm',
        'C:\Program Files (x86)\MathType\Office Support\32\MathType Commands 2016.dotm',
        'C:\Program Files\MathType\Office Support\32\MathType Commands 2016.dotm'
    )
    $mathTypeTemplate = $mathTypeTemplateCandidates | Where-Object {
        Test-Path -LiteralPath $_ -PathType Leaf
    } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($mathTypeTemplate)) {
        throw 'MathType Commands template was not found. Verify that MathType is installed for Microsoft Word.'
    }

    try {
        $addIns = $word.AddIns
        $mathTypeAddIn = $addIns.Add($mathTypeTemplate, $true)
        $mathTypeAddIn.Installed = $true
        Write-Log "MathType commands loaded: $mathTypeTemplate"
    }
    catch {
        throw "Cannot load the MathType commands template. Details: $($_.Exception.Message)"
    }

    try {
        $documents = $word.Documents
        $document = $documents.Open($OutputPath, $false, $false, $false)
        $window = $document.ActiveWindow
        $nativeProcessId = [uint32]0
        [void][WordWindowNative]::GetWindowThreadProcessId(
            [IntPtr]([int64]$window.Hwnd),
            [ref]$nativeProcessId
        )
        $wordProcessId = [int]$nativeProcessId
        Activate-WordWindow
    }
    catch {
        throw "Cannot open the working copy in Word: $OutputPath. Details: $($_.Exception.Message)"
    }

    # StoryRanges covers the main story (including tables), headers, footers,
    # footnotes/endnotes/comments and text-frame stories exposed safely by Word.
    $storyRanges = $document.StoryRanges
    $storyInstance = 0
    for ($storyType = 1; $storyType -le 17; $storyType++) {
        $story = $null
        try { $story = $storyRanges.Item($storyType) } catch { $story = $null }

        while ($null -ne $story) {
            $storyInstance++
            $storyText = [string]$story.Text
            $openAt = -1

            for ($offset = 0; $offset -lt $storyText.Length; $offset++) {
                if ($storyText[$offset] -ne '$') { continue }

                if ($openAt -lt 0) {
                    $openAt = $offset
                    continue
                }

                # Do not join delimiters across paragraph/cell/story boundaries.
                $inside = $storyText.Substring($openAt + 1, $offset - $openAt - 1)
                if ($inside.IndexOfAny([char[]]@("`r", "`n", [char]7)) -ge 0) {
                    $openAt = $offset
                    continue
                }

                $range = $story.Duplicate
                $range.SetRange($story.Start + $openAt, $story.Start + $offset + 1)
                [void]$equations.Add([pscustomobject]@{
                    Range         = $range
                    StoryType     = $storyType
                    StoryInstance = $storyInstance
                    Start         = [int]$range.Start
                    Text          = [string]$range.Text
                })
                $openAt = -1
            }

            $nextStory = $null
            try { $nextStory = $story.NextStoryRange } catch { $nextStory = $null }
            Release-ComObject $story
            $story = $nextStory
        }
    }

    $orderedEquations = @($equations | Sort-Object StoryInstance, @{ Expression = 'Start'; Descending = $true })
    Write-Log "Equations found: $($orderedEquations.Count)"

    for ($index = 0; $index -lt $orderedEquations.Count; $index++) {
        $item = $orderedEquations[$index]
        $range = $item.Range
        $preview = Get-Preview $item.Text
        Write-Log ("Processing {0}/{1} (story {2}): {3}" -f ($index + 1), $orderedEquations.Count, $item.StoryType, $preview)

        $converted = $false
        $lastMessage = $null
        for ($attempt = 1; $attempt -le 2 -and -not $converted; $attempt++) {
            try {
                Activate-WordWindow
                $range.Select()
                Activate-WordWindow

                # This is the Ribbon callback behind MathType's Alt+\ Toggle TeX command.
                # A null IRibbonControl is accepted by the signed MathType template.
                $ribbonControl = $null
                $word.Run('MTCommand_OnTexToggle', [ref]$ribbonControl)
                Start-Sleep -Milliseconds $WaitMilliseconds
                $converted = Test-RangeConverted $range $item.Text

                if (-not $converted) {
                    $lastMessage = "Toggle TeX did not replace the text on attempt $attempt"
                    if ($attempt -eq 1) {
                        Write-Log "Not converted; retrying once." 'WARN'
                        Start-Sleep -Milliseconds 250
                    }
                }
            }
            catch {
                $lastMessage = $_.Exception.Message
                if ($attempt -eq 1) {
                    Write-Log "First attempt failed; retrying once: $lastMessage" 'WARN'
                    Start-Sleep -Milliseconds 250
                }
            }
        }

        # Try a narrowly targeted fallback only after the two normal Toggle TeX
        # attempts have failed.
        if (-not $converted) {
            $fallback = Get-DollarTeXFallback -Text $item.Text

            if ($null -ne $fallback) {
                Write-Log (
                    "Trying {0} fallback: {1}" -f
                    $fallback.Category,
                    $fallback.FallbackText
                ) 'WARN'

                $fallbackResult = Invoke-DollarTeXFallback `
                    -Range $range `
                    -OriginalText $item.Text `
                    -Fallback $fallback `
                    -WaitMilliseconds $WaitMilliseconds

                if ($fallbackResult.Attempted) {
                    $converted = [bool]$fallbackResult.Converted
                    $lastMessage = [string]$fallbackResult.Message

                    if ($converted) {
                        Write-Log $lastMessage 'INFO'
                    }
                    else {
                        Write-Log $lastMessage 'WARN'
                    }
                }
            }
        }

        if ($converted) {
            $successCount++
            Write-Log "Succeeded: $successCount; failed: $failureCount"
        }
        else {
            $failureCount++
            Write-Log "Failed for '$preview'. Reason: $lastMessage" 'ERROR'
        }
    }

    $document.Save()
    $savedOutput = $true
    Write-Log "Succeeded: $successCount"
    Write-Log "Failed: $failureCount"
    Write-Log "Output file: $OutputPath"

    if ($failureCount -gt 0) {
        throw "MathType Toggle TeX failed for $failureCount/$($orderedEquations.Count) equations. The working copy was saved for inspection; the input file was not changed."
    }
}
catch {
    $fatalError = $_
    Write-Log $_.Exception.Message 'ERROR'
}
finally {
    if ($null -ne $document) {
        if (-not $savedOutput -and $createdOutput) {
            try { $document.Save() } catch {}
        }
        try { $document.Close($wdDoNotSaveChanges) } catch {}
    }
    if ($null -ne $word) {
        try { $word.Quit($wdDoNotSaveChanges) } catch {}
    }

    for ($i = 0; $i -lt $equations.Count; $i++) {
        Release-ComObject $equations[$i].Range
    }
    Release-ComObject $storyRanges
    Release-ComObject $window
    Release-ComObject $document
    Release-ComObject $documents
    Release-ComObject $mathTypeAddIn
    Release-ComObject $addIns
    Release-ComObject $shell
    Release-ComObject $word

    $orderedEquations = $null
    $equations = $null
    $storyRanges = $null
    $window = $null
    $document = $null
    $documents = $null
    $mathTypeAddIn = $null
    $addIns = $null
    $shell = $null
    $word = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Some MathType versions keep an empty /Automation Word process alive even
    # after Application.Quit. Only clean up the exact process created here.
    if ($wordProcessId -gt 0) {
        $wordProcess = $null
        try {
            $wordProcess = [Diagnostics.Process]::GetProcessById($wordProcessId)
            if (-not $wordProcess.WaitForExit(4000)) {
                [void]$wordProcess.CloseMainWindow()
                if (-not $wordProcess.WaitForExit(2000)) {
                    Write-Log "Stopping the empty Word automation process (PID $wordProcessId)." 'WARN'
                    $wordProcess.Kill()
                    [void]$wordProcess.WaitForExit(3000)
                }
            }
        }
        catch [ArgumentException] {
            # The process already exited normally.
        }
        finally {
            if ($null -ne $wordProcess) { $wordProcess.Dispose() }
        }
    }
}

if ($null -ne $fatalError) {
    throw $fatalError
}

```

# 2. Auto-download

Tool: [`gdrive-videoloader`](FILE\gdrive-videoloadeo.zip).

Syntax:

```sh
powershell -ExecutionPolicy Bypass -File "F:\gia-su\video-driver\gdrive-videoloader-main\gdrive-videoloader-main\run.ps1"
```

```sh
# auto_gdrive.ps1
# Yeu cau: PowerShell 7+ (chay bang lenh pwsh, khong dung powershell.exe)
# Tu dong: mo Chrome debug (neu chua mo) -> F5 Drive -> lay cookie -> chay download lan luot
# Video tai ve se luu vao D:\VIDEO-KHOA-HOC, giu nguyen ten goc tren Google Drive

$ErrorActionPreference = "Stop"

$workDir       = "F:\gia-su\video-driver\gdrive-videoloader-main\gdrive-videoloader-main"
$urlFile       = Join-Path $workDir "url.txt"
$cookiesPath   = Join-Path $workDir "cookies.txt"
$driveFolder   = "https://drive.google.com/drive/folders/10WZDzp6DFRAdiyZhb9k3h-ApbxkYEpLd"
$cdpBase       = "http://127.0.0.1:9222"   # dung 127.0.0.1 thay vi localhost de tranh loi phan giai IPv6 (::1)

$chromeExe     = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$chromeProfile = "F:\gia-su\chrome-automation-profile"
$chromePort    = 9222

$outputDir     = "D:\VIDEO-KHOA-HOC"   # noi video se duoc luu

function Test-CdpAlive {
    param([switch]$Verbose)
    try {
        Invoke-RestMethod "$cdpBase/json/version" -TimeoutSec 2 | Out-Null
        return $true
    } catch {
        if ($Verbose) {
            Write-Host "     [debug] Loi ket noi CDP: $($_.Exception.Message)" -ForegroundColor Red
        }
        return $false
    }
}

function Start-ChromeDebugIfNeeded {
    if (Test-CdpAlive) {
        Write-Host "-> Chrome debug da chay san, dung lai." -ForegroundColor DarkGray
        return
    }

    # Chi coi la "da chay" neu tim thay dung TIEN TRINH CHINH cua Chrome
    # (co flag --remote-debugging-port, KHONG phai tien trinh con nhu renderer/GPU
    # vi cac tien trinh con cung ten chrome.exe va cung ke thua --user-data-dir).
    $mainProc = Get-CimInstance Win32_Process -Filter "Name = 'chrome.exe'" -ErrorAction SilentlyContinue |
        Where-Object {
            $_.CommandLine -and
            $_.CommandLine.Contains($chromeProfile) -and
            $_.CommandLine.Contains("--remote-debugging-port") -and
            -not $_.CommandLine.Contains("--type=")
        }

    if ($mainProc) {
        Write-Host "-> Phat hien Chrome dang khoi dong voi profile nay, doi no san sang thay vi mo moi..." -ForegroundColor Yellow
    } else {
        # Don sach cac tien trinh chrome.exe mo coi (con/renderer sot lai) dung profile nay,
        # neu khong chung co the giu Singleton lock khien Chrome moi khong khoi dong duoc.
        $orphans = Get-CimInstance Win32_Process -Filter "Name = 'chrome.exe'" -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandLine -and $_.CommandLine.Contains($chromeProfile) }
        if ($orphans) {
            Write-Host "-> Don cac tien trinh Chrome mo coi cua profile nay truoc khi mo lai..." -ForegroundColor DarkYellow
            $orphans | ForEach-Object {
                Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
            }
            Start-Sleep -Seconds 1
            Get-ChildItem -Path $chromeProfile -Filter "Singleton*" -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }

        Write-Host "-> Chua thay Chrome debug, dang mo moi..." -ForegroundColor Yellow
        $proc = Start-Process -FilePath $chromeExe -ArgumentList @(
            "--remote-debugging-port=$chromePort",
            "--user-data-dir=$chromeProfile",
            "--no-first-run",
            "--no-default-browser-check",
            "--remote-allow-origins=*"
        ) -PassThru
    }

    $maxWait = 60   # tang thoi gian cho vi Chrome co the khoi dong cham (o F: cham, antivirus quet profile moi, v.v.)
    $waited = 0
    while (-not (Test-CdpAlive -Verbose:($waited % 5 -eq 0 -and $waited -gt 0))) {
        Start-Sleep -Seconds 1
        $waited++
        if ($waited % 5 -eq 0) {
            Write-Host "  -> Van dang doi Chrome debug san sang... ($waited/$maxWait giay)" -ForegroundColor DarkGray
            if ($proc -and $proc.HasExited) {
                throw "Tien trinh Chrome vua mo da THOAT (exit code $($proc.ExitCode)) truoc khi debug port san sang. Kiem tra profile $chromeProfile co bi loi/khoa khong, hoac chay thu chrome.exe thu cong voi cung tham so de xem loi."
            }
        }
        if ($waited -ge $maxWait) {
            throw "Chrome debug khong khoi dong duoc sau $maxWait giay. Kiem tra: 1) duong dan chrome.exe, 2) firewall/antivirus co chan 127.0.0.1:$chromePort khong (thu tat AV tam thoi), 3) o dia chua $chromeProfile co dang cham/bi treo khong, 4) mo thu trinh duyet va truy cap http://127.0.0.1:$chromePort/json/version xem co tra ve JSON khong."
        }
    }
    Write-Host "-> Chrome debug da san sang (mat $waited giay)." -ForegroundColor DarkGray
}

function Get-CdpTab {
    $tabs = Invoke-RestMethod "$cdpBase/json"
    $tab = $tabs | Where-Object { $_.type -eq "page" } | Select-Object -First 1
    if (-not $tab) {
        $tab = Invoke-RestMethod "$cdpBase/json/new?$driveFolder" -Method PUT
    }
    return $tab
}

function Invoke-CdpCommand {
    param(
        [string]$WsUrl,
        [string]$Method,
        [hashtable]$Params = @{}
    )
    $ws = [System.Net.WebSockets.ClientWebSocket]::new()
    $uri = [Uri]$WsUrl
    $ws.ConnectAsync($uri, [Threading.CancellationToken]::None).Wait()

    $id = Get-Random -Maximum 999999
    $payload = @{ id = $id; method = $Method; params = $Params } | ConvertTo-Json -Depth 10 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
    $seg = [ArraySegment[byte]]::new($bytes)
    $ws.SendAsync($seg, [Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).Wait()

    $buffer = [byte[]]::new(1MB)
    $result = ""
    do {
        $recvSeg = [ArraySegment[byte]]::new($buffer)
        $recv = $ws.ReceiveAsync($recvSeg, [Threading.CancellationToken]::None).Result
        $result += [Text.Encoding]::UTF8.GetString($buffer, 0, $recv.Count)
    } while (-not $recv.EndOfMessage)

    $ws.CloseAsync([Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "done", [Threading.CancellationToken]::None).Wait()
    return $result | ConvertFrom-Json
}

function Update-CookiesFromChrome {
    Write-Host "  -> Dang F5 trang Google Drive va lay cookie moi..." -ForegroundColor DarkCyan

    $tab = Get-CdpTab
    $wsUrl = $tab.webSocketDebuggerUrl

    Invoke-CdpCommand -WsUrl $wsUrl -Method "Page.navigate" -Params @{ url = $driveFolder } | Out-Null
    Start-Sleep -Seconds 6

    $cookieResp = Invoke-CdpCommand -WsUrl $wsUrl -Method "Network.getAllCookies"
    $cookies = $cookieResp.result.cookies

    if (-not $cookies -or $cookies.Count -eq 0) {
        throw "Khong lay duoc cookie nao - kiem tra lai da dang nhap Google trong Chrome debug profile chua."
    }

    $lines = @("# Netscape HTTP Cookie File")
    foreach ($c in $cookies) {
        if ($c.domain -notmatch "google\.com$") { continue }
        $includeSub = if ($c.domain.StartsWith(".")) { "TRUE" } else { "FALSE" }
        $secure = if ($c.secure) { "TRUE" } else { "FALSE" }
        $expiry = if ($c.expires -and $c.expires -gt 0) { [int64]$c.expires } else { 0 }
        $lines += "$($c.domain)`t$includeSub`t$($c.path)`t$secure`t$expiry`t$($c.name)`t$($c.value)"
    }
    Set-Content -Path $cookiesPath -Value $lines -Encoding UTF8
    Write-Host "  -> Da luu cookie moi vao $cookiesPath" -ForegroundColor DarkCyan
}

# ---- Main ----
Start-ChromeDebugIfNeeded

if (-not (Test-Path $outputDir)) {
    Write-Host "-> Thu muc $outputDir chua ton tai, dang tao..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$urls = Get-Content $urlFile | Where-Object { $_.Trim() -ne "" }
$total = $urls.Count
$i = 0

foreach ($url in $urls) {
    $i++
    Write-Host "=== Bat dau $i/$total : $url ===" -ForegroundColor Cyan

    Update-CookiesFromChrome

    # Chuyen sang thu muc luu video truoc khi chay python
    Set-Location $outputDir
    python (Join-Path $workDir "gdrive_videoloader.py") $url --cookies "$cookiesPath"

    Write-Host "=== Chay xong $i/$total ===" -ForegroundColor Green
}

Write-Host "Hoan tat toan bo $total url. Video da luu tai $outputDir" -ForegroundColor Yellow
```