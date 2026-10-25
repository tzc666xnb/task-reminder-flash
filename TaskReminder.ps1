#requires -Version 5.1
<#
    任务提醒 TaskReminder
    - 托盘常驻，到点弹窗 + 提示音
    - 任务可循环（每天/每周/每月/工作日），完成后自动排下一次
    - 随时添加任务、勾选完成、稍后提醒
    - 开机自启动（写入启动文件夹快捷方式）
    - 数据保存在 data\tasks.json，纯本地、不联网

    用法：
        powershell -ExecutionPolicy Bypass -File TaskReminder.ps1
        powershell -ExecutionPolicy Bypass -File TaskReminder.ps1 -Setup           # 建桌面 + 开机启动快捷方式
        powershell -ExecutionPolicy Bypass -File TaskReminder.ps1 -RemoveShortcuts
        powershell -ExecutionPolicy Bypass -File TaskReminder.ps1 -SelfTest        # 逻辑自检
#>
param(
    [switch]$SelfTest,
    [switch]$Setup,
    [switch]$RemoveShortcuts,
    [switch]$Preview,
    [switch]$StartHidden
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Core

$script:AppDir   = Split-Path -Parent $MyInvocation.MyCommand.Definition
$script:DataDir  = Join-Path $script:AppDir 'data'
$script:DataFile = Join-Path $script:DataDir 'tasks.json'
$script:SetFile  = Join-Path $script:DataDir 'settings.json'
$script:IconPath = Join-Path $script:AppDir 'app.ico'
$script:AppName  = '任务提醒'
$script:Version  = '1.4'

$script:Accent = [System.Drawing.Color]::FromArgb(59, 110, 246)
$script:Ink    = [System.Drawing.Color]::FromArgb(24, 33, 48)
$script:Ink2   = [System.Drawing.Color]::FromArgb(107, 118, 136)
$script:Line   = [System.Drawing.Color]::FromArgb(232, 236, 242)
$script:Bg     = [System.Drawing.Color]::White
$script:SoftBg = [System.Drawing.Color]::FromArgb(248, 250, 253)
$script:Danger = [System.Drawing.Color]::FromArgb(214, 69, 69)

$script:UiFont  = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5)
$script:UiBold  = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5, [System.Drawing.FontStyle]::Bold)
$script:BigFont = New-Object System.Drawing.Font('Microsoft YaHei UI', 15, [System.Drawing.FontStyle]::Bold)

$script:RepeatKeys   = @('none', 'daily', 'weekly', 'monthly', 'weekday')
$script:RepeatLabels = @('不重复', '每天', '每周', '每月', '工作日')

$script:Tasks    = New-Object System.Collections.ArrayList
$script:Settings = [pscustomobject]@{ Sound = $true; MinimizeToTray = $true; Launched = 0; NagMinutes = 15 }
$script:Filter   = 'todo'
$script:Updating = $false
$script:Exiting  = $false
$script:Flash    = $null
$script:ToastForms = New-Object System.Collections.ArrayList

# ================================================================ 基础工具

function Write-Log {
    param([string]$Message)
    try {
        $f = Join-Path $script:AppDir 'data\app.log'
        if ((Test-Path $f) -and ((Get-Item $f).Length -gt 300000)) { Remove-Item $f -Force }
        Add-Content -LiteralPath $f -Value ("{0}  {1}" -f (Get-Date -Format 'MM-dd HH:mm:ss'), $Message) -Encoding UTF8
    } catch { }
}

function ConvertTo-DateTime {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    try { return [datetime]::ParseExact($Text, 'yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) }
    catch { return $null }
}

function Get-TaskDeadline {
    param($Task)
    $e = ConvertTo-DateTime $Task.end
    if ($e) { return $e }
    return (ConvertTo-DateTime $Task.due)
}

function Get-AddSlot {
    $n = Get-Date
    if ($n.Minute -lt 30) { return $n.Date.AddHours($n.Hour).AddMinutes(30) }
    return $n.Date.AddHours($n.Hour).AddMinutes(60)
}

function Get-DueText {

    param([datetime]$Due, [bool]$Done, $End)
    $now = Get-Date
    $d = $Due.Date
    if ($d -eq $now.Date) { $label = '今天' }
    elseif ($d -eq $now.Date.AddDays(1)) { $label = '明天' }
    elseif ($d -eq $now.Date.AddDays(-1)) { $label = '昨天' }
    else { $label = $Due.ToString('MM-dd') }

    $text = "$label $($Due.ToString('HH:mm'))"
    $deadline = $Due
    if ($End) {
        $e = [datetime]$End
        if ($e.Date -eq $Due.Date) { $text = $text + '-' + $e.ToString('HH:mm') }
        else { $text = $text + '→次日' + $e.ToString('HH:mm') }
        $deadline = $e
    }

    if (-not $Done) {
        if ($now -gt $deadline) { $text = $text + ' 逾期' }
        elseif ($End -and $now -ge $Due) { $text = $text + ' 进行中' }
    }
    return $text
}

function Get-RepeatLabel {
    param([string]$Repeat)
    $i = $script:RepeatKeys.IndexOf([string]$Repeat)
    if ($i -lt 0) { return '不重复' }
    return $script:RepeatLabels[$i]
}

function Get-NextDue {
    param([datetime]$From, [string]$Repeat)
    switch ($Repeat) {
        'daily'   { return $From.AddDays(1) }
        'weekly'  { return $From.AddDays(7) }
        'monthly' { return $From.AddMonths(1) }
        'weekday' {
            $n = $From.AddDays(1)
            while ($n.DayOfWeek -eq 'Saturday' -or $n.DayOfWeek -eq 'Sunday') { $n = $n.AddDays(1) }
            return $n
        }
        default { return $From.AddDays(1) }
    }
}

function Get-DueTasks {
    param([datetime]$Now)
    $out = New-Object System.Collections.ArrayList
    $nag = 15
    if ($null -ne $script:Settings.NagMinutes) { $nag = [int]$script:Settings.NagMinutes }
    foreach ($t in $script:Tasks) {
        if ([bool]$t.done) { continue }
        $eff = ConvertTo-DateTime $t.due
        $snoozed = $false
        if ($t.snoozeUntil) {
            $s = ConvertTo-DateTime $t.snoozeUntil
            if ($s) { $eff = $s; $snoozed = $true }
        }
        if (-not $eff) { continue }

        if ($snoozed) { $remind = $eff }
        else { $remind = $eff.AddMinutes(-1 * [int]$t.lead) }
        if ($Now -lt $remind) { continue }

        if (-not [bool]$t.notified) { [void]$out.Add($t); continue }

        # 已经提醒过：等过了结束时间，并且超过间隔，再催一次
        if ($nag -le 0) { continue }
        $deadline = Get-TaskDeadline $t
        if (-not $deadline) { $deadline = $eff }
        if ($Now -lt $deadline) { continue }
        $last = ConvertTo-DateTime $t.lastAlert
        if (-not $last) { $last = $eff }
        if (($Now - $last).TotalMinutes -ge $nag) { [void]$out.Add($t) }
    }
    return ,$out
}

function Get-OverdueText {
    param([datetime]$Due)
    $span = (Get-Date) - $Due
    if ($span.TotalMinutes -lt 60) {
        $m = [int]$span.TotalMinutes
        if ($m -lt 1) { $m = 1 }
        return "$m 分钟"
    }
    if ($span.TotalHours -lt 24) { return ("{0} 小时" -f [int]$span.TotalHours) }
    return ("{0} 天" -f [int]$span.TotalDays)
}

function Mark-Alerted {
    param($Task)
    if (-not $Task) { return }
    $Task.notified = $true
    $Task | Add-Member -NotePropertyName lastAlert -NotePropertyValue ((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')) -Force
}

function Get-TaskById {
    param([string]$Id)
    foreach ($t in $script:Tasks) { if ($t.id -eq $Id) { return $t } }
    return $null
}

# ================================================================ 存储

function Move-TaskById {
    param([string]$Id, [datetime]$To)
    $t = Get-TaskById $Id
    if (-not $t) { return }
    $oldDue = ConvertTo-DateTime $t.due
    $delta = [TimeSpan]::Zero
    if ($oldDue) { $delta = $To - $oldDue }
    $t.due = $To.ToString('yyyy-MM-ddTHH:mm:ss')
    $t.notified = $false
    $t.snoozeUntil = $null
    $t | Add-Member -NotePropertyName lastAlert -NotePropertyValue $null -Force
    if ([bool]$t.done) { $t.done = $false; $t.doneAt = $null }
    $ed = ConvertTo-DateTime $t.end
    if ($ed -and $delta -ne [TimeSpan]::Zero) {
        $t | Add-Member -NotePropertyName end -NotePropertyValue ($ed.Add($delta).ToString('yyyy-MM-ddTHH:mm:ss')) -Force
    }
    Save-Tasks
}

function Initialize-FileApi {
    if ($null -ne ('TrFs.Native' -as [type])) { return }
    try {
        Add-Type -Namespace TrFs -Name Native -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
public static extern bool MoveFileEx(string existingFileName, string newFileName, int flags);
"@
    } catch { }
}

function Save-Tasks {
    if (-not (Test-Path $script:DataDir)) { New-Item -ItemType Directory -Path $script:DataDir -Force | Out-Null }
    Initialize-FileApi
    $json = ConvertTo-Json -InputObject @($script:Tasks) -Depth 6
    $tmp = $script:DataFile + '.tmp'
    $bak = $script:DataFile + '.old'
    $err = $null
    for ($i = 1; $i -le 5; $i++) {
        try {
            [System.IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding($false)))
            if (Test-Path $script:DataFile) {
                try { Copy-Item -LiteralPath $script:DataFile -Destination $bak -Force -ErrorAction Stop } catch { }
            }
            # 单次 rename（MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH），不存在"目标先消失"的窗口
            $moved = $false
            try { $moved = [TrFs.Native]::MoveFileEx($tmp, $script:DataFile, 9) } catch { $moved = $false }
            if (-not $moved) { throw 'MoveFileEx failed' }
            if (Test-Path $bak) { Remove-Item -LiteralPath $bak -Force -ErrorAction SilentlyContinue }
            return $true
        } catch {
            $err = $_.Exception.Message
            Start-Sleep -Milliseconds (80 * $i)
        }
    }
    Write-Log ('save failed: ' + $err)
    if (Test-Path $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    $script:Flash = '保存失败，请重试（文件可能被占用）'
    return $false
}

function Load-Tasks {
    $list = New-Object System.Collections.ArrayList
    if (-not (Test-Path $script:DataFile)) {
        $stale = $script:DataFile + '.old'
        if (Test-Path $stale) {
            Copy-Item -LiteralPath $stale -Destination $script:DataFile -Force -ErrorAction SilentlyContinue
            Write-Log 'recovered tasks.json from .old'
        }
    }
    if (Test-Path $script:DataFile) {
        try {
            $raw = Get-Content -LiteralPath $script:DataFile -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $data = ConvertFrom-Json -InputObject $raw
                if ($null -ne $data -and $null -ne $data.PSObject.Properties['value'] -and $null -ne $data.PSObject.Properties['Count']) {
                    $data = $data.value
                }
                $dropped = 0
                foreach ($t in @($data)) {
                    if ($null -eq $t) { $dropped++; continue }
                    if ($null -eq $t.PSObject.Properties['title']) { $dropped++; continue }
                    if ($null -eq $t.PSObject.Properties['due']) { $dropped++; continue }
                    [void]$list.Add($t)
                }
                if ($dropped -gt 0) {
                    $bak = $script:DataFile + '.broken-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
                    Copy-Item -LiteralPath $script:DataFile -Destination $bak -Force -ErrorAction SilentlyContinue
                    Write-Log "load dropped $dropped malformed entries, backup: $bak"
                }
            }
        } catch {
            $bak = $script:DataFile + '.broken-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
            Copy-Item -LiteralPath $script:DataFile -Destination $bak -Force -ErrorAction SilentlyContinue
        }
    }
    return ,$list
}

function Save-Settings {
    if (-not (Test-Path $script:DataDir)) { New-Item -ItemType Directory -Path $script:DataDir -Force | Out-Null }
    $json = ConvertTo-Json -InputObject $script:Settings -Depth 4
    for ($i = 1; $i -le 3; $i++) {
        try {
            [System.IO.File]::WriteAllText($script:SetFile, $json, (New-Object System.Text.UTF8Encoding($false)))
            return $true
        } catch {
            Start-Sleep -Milliseconds (60 * $i)
        }
    }
    return $false
}

function Load-Settings {
    if (Test-Path $script:SetFile) {
        try {
            $s = Get-Content -LiteralPath $script:SetFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -ne $s.Sound) { $script:Settings.Sound = [bool]$s.Sound }
            if ($null -ne $s.MinimizeToTray) { $script:Settings.MinimizeToTray = [bool]$s.MinimizeToTray }
            if ($null -ne $s.Launched) { $script:Settings.Launched = [int]$s.Launched }
            if ($null -ne $s.NagMinutes) { $script:Settings.NagMinutes = [int]$s.NagMinutes }
        } catch { }
    }
}

# ================================================================ 图标 / 快捷方式

function New-AppIcon {
    param([string]$Path)
    $size = 256
    $bmp = New-Object System.Drawing.Bitmap($size, $size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)

    $rect = New-Object System.Drawing.Rectangle(8, 8, 240, 240)
    $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
    $r = 56
    $gp.AddArc($rect.X, $rect.Y, $r, $r, 180, 90)
    $gp.AddArc($rect.Right - $r, $rect.Y, $r, $r, 270, 90)
    $gp.AddArc($rect.Right - $r, $rect.Bottom - $r, $r, $r, 0, 90)
    $gp.AddArc($rect.X, $rect.Bottom - $r, $r, $r, 90, 90)
    $gp.CloseFigure()

    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, [System.Drawing.Color]::FromArgb(96, 150, 255), [System.Drawing.Color]::FromArgb(43, 84, 214), 55)
    $g.FillPath($brush, $gp)

    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 24)
    $pen.StartCap = 'Round'; $pen.EndCap = 'Round'; $pen.LineJoin = 'Round'
    $pts = @(
        (New-Object System.Drawing.Point(76, 132)),
        (New-Object System.Drawing.Point(112, 168)),
        (New-Object System.Drawing.Point(184, 88))
    )
    $g.DrawLines($pen, $pts)
    $g.Dispose()

    $pngTmp = Join-Path $env:TEMP ('tr_icon_' + [guid]::NewGuid().ToString('N') + '.png')
    $bmp.Save($pngTmp, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $png = [System.IO.File]::ReadAllBytes($pngTmp)
    Remove-Item $pngTmp -Force -ErrorAction SilentlyContinue

    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)
    $bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]1)
    $bw.Write([byte]0);  $bw.Write([byte]0)          # 0 = 256px
    $bw.Write([byte]0);  $bw.Write([byte]0)          # palette, reserved
    $bw.Write([uint16]1); $bw.Write([uint16]32)      # planes, bpp
    $bw.Write([uint32]$png.Length)
    $bw.Write([uint32]22)
    $bw.Write($png)
    $bw.Flush()
    [System.IO.File]::WriteAllBytes($Path, $ms.ToArray())
    $bw.Dispose(); $ms.Dispose()
}

function Get-StartupShortcut { Join-Path ([Environment]::GetFolderPath('Startup')) ($script:AppName + '.lnk') }
function Get-DesktopShortcut { Join-Path ([Environment]::GetFolderPath('Desktop')) ($script:AppName + '.lnk') }
function Test-Autostart { return (Test-Path (Get-StartupShortcut)) }

function New-Shortcut {
    param([string]$LinkPath, [string]$VbsPath, [string]$ExtraArgs = '', [int]$WindowStyle = 1)
    $ws = New-Object -ComObject WScript.Shell
    $sc = $ws.CreateShortcut($LinkPath)
    $sc.TargetPath = Join-Path $env:SystemRoot 'System32\wscript.exe'
    if ($ExtraArgs) { $sc.Arguments = '"' + $VbsPath + '" ' + $ExtraArgs } else { $sc.Arguments = '"' + $VbsPath + '"' }
    $sc.WorkingDirectory = $script:AppDir
    $sc.WindowStyle = $WindowStyle
    if (Test-Path $script:IconPath) { $sc.IconLocation = $script:IconPath + ',0' }
    $sc.Description = $script:AppName
    $sc.Save()
}

function Set-Autostart {
    param([bool]$Enabled)
    $lnk = Get-StartupShortcut
    if ($Enabled) {
        if (-not (Test-Path $script:IconPath)) { New-AppIcon -Path $script:IconPath }
        New-Shortcut -LinkPath $lnk -VbsPath (Join-Path $script:AppDir 'Start-TaskReminder.vbs') -ExtraArgs '-StartHidden' -WindowStyle 1
    } elseif (Test-Path $lnk) {
        Remove-Item -LiteralPath $lnk -Force
    }
    return (Test-Path $lnk)
}

# ================================================================ 任务操作

function Complete-TaskById {
    param([string]$Id)
    $t = Get-TaskById $Id
    if (-not $t) { return }
    $now = Get-Date
    if ($t.repeat -and [string]$t.repeat -ne 'none') {
        $base = ConvertTo-DateTime $t.due
        if (-not $base) { $base = $now }
        $next = Get-NextDue -From $base -Repeat ([string]$t.repeat)
        $guard = 0
        while ($next -le $now -and $guard -lt 5000) {
            $next = Get-NextDue -From $next -Repeat ([string]$t.repeat)
            $guard++
        }
        if ($next -le $now) { $next = $now.AddMinutes(1) }
        $t.due = $next.ToString('yyyy-MM-ddTHH:mm:ss')
        # 时段任务：结束时间跟着整体平移
        $ed = ConvertTo-DateTime $t.end
        if ($ed) {
            $shifted = $ed.Add($next - $base)
            $t | Add-Member -NotePropertyName end -NotePropertyValue ($shifted.ToString('yyyy-MM-ddTHH:mm:ss')) -Force
        }
        $t.notified = $false
        $t.snoozeUntil = $null
        $t | Add-Member -NotePropertyName lastAlert -NotePropertyValue $null -Force
        $t.completedCount = [int]$t.completedCount + 1
        $script:Flash = "已完成 · 下次 $($next.ToString('MM-dd HH:mm'))"
    } else {
        $t.done = $true
        $t.doneAt = $now.ToString('yyyy-MM-ddTHH:mm:ss')
        $t.notified = $true
        $t.snoozeUntil = $null
        $script:Flash = '已完成'
    }
    Save-Tasks
}

function Snooze-TaskById {
    param([string]$Id, [int]$Minutes)
    $t = Get-TaskById $Id
    if (-not $t) { return }
    $t.snoozeUntil = (Get-Date).AddMinutes($Minutes).ToString('yyyy-MM-ddTHH:mm:ss')
    $t.notified = $false
    $t | Add-Member -NotePropertyName lastAlert -NotePropertyValue $null -Force
    if ([bool]$t.done) { $t.done = $false; $t.doneAt = $null }
    $script:Flash = "$Minutes 分钟后再次提醒"
    Save-Tasks
}

function Remove-TaskById {
    param([string]$Id)
    $t = Get-TaskById $Id
    if (-not $t) { return }
    [void]$script:Tasks.Remove($t)
    Save-Tasks
}

# ================================================================ 提醒弹窗

function Reposition-Toasts {
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $i = 0
    foreach ($f in @($script:ToastForms)) {
        if ($f.IsDisposed -or -not $f.Visible) { continue }
        $f.Location = New-Object System.Drawing.Point(($wa.Right - $f.Width - 20), ($wa.Bottom - $f.Height - 20 - ($i * ($f.Height + 10))))
        $i++
    }
}

function Show-Toast {
    param($Task)

    $f = New-Object System.Windows.Forms.Form
    $f.FormBorderStyle = 'None'
    $f.StartPosition = 'Manual'
    $f.Size = New-Object System.Drawing.Size(392, 152)
    $f.BackColor = [System.Drawing.Color]::White
    $f.TopMost = $true
    $f.ShowInTaskbar = $false
    $f.Tag = [string]$Task.id

    $bar = New-Object System.Windows.Forms.Panel
    $bar.Dock = 'Left'; $bar.Width = 5; $bar.BackColor = $script:Accent
    $f.Controls.Add($bar)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Text = '任务提醒'
    $lblHead.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 8.5)
    $lblHead.ForeColor = $script:Ink2
    $lblHead.AutoSize = $false
    $lblHead.SetBounds(20, 12, 300, 18)
    $lblHead.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($lblHead)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = [string]$Task.title
    $lblTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $script:Ink
    $lblTitle.AutoSize = $false
    $lblTitle.AutoEllipsis = $true
    $lblTitle.SetBounds(20, 32, 350, 26)
    $lblTitle.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($lblTitle)

    $dueDt = ConvertTo-DateTime $Task.due
    $snoozeDt = $null
    if ($Task.snoozeUntil) { $snoozeDt = ConvertTo-DateTime $Task.snoozeUntil }
    $eff = $dueDt
    if ($snoozeDt) { $eff = $snoozeDt }
    $endDt = ConvertTo-DateTime $Task.end
    if (-not $dueDt) { $dueDt = $snoozeDt }
    if ($snoozeDt -and $dueDt) {
        $sub = '稍后提醒 ' + $snoozeDt.ToString('HH:mm') + ' · 原定 ' + $dueDt.ToString('HH:mm')
    } elseif ($endDt) {
        $sub = '时段 ' + $eff.ToString('MM-dd HH:mm') + ' → ' + $endDt.ToString('HH:mm')
    } else {
        $sub = '截止 ' + $eff.ToString('MM-dd HH:mm')
    }
    $deadline = $endDt
    if (-not $deadline) { $deadline = $dueDt }
    if ((Get-Date) -gt $deadline) { $sub = $sub + ' · 已逾期 ' + (Get-OverdueText $deadline) }
    elseif ($endDt -and -not $snoozeDt -and (Get-Date) -ge $dueDt) { $sub = $sub + ' · 进行中' }
    if ($Task.repeat -and [string]$Task.repeat -ne 'none') { $sub = $sub + ' · ' + (Get-RepeatLabel ([string]$Task.repeat)) }
    if ($Task.note) { $sub = $sub + ' · ' + [string]$Task.note }

    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text = $sub
    $lblSub.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 8.5)
    $lblSub.ForeColor = $script:Ink2
    $lblSub.AutoSize = $false
    $lblSub.AutoEllipsis = $true
    $lblSub.SetBounds(20, 58, 350, 18)
    $lblSub.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($lblSub)

    $btnDone = New-Object System.Windows.Forms.Button
    $btnDone.Text = '完 成'
    $btnDone.Font = $script:UiBold
    $btnDone.FlatStyle = 'Flat'
    $btnDone.FlatAppearance.BorderSize = 0
    $btnDone.BackColor = $script:Accent
    $btnDone.ForeColor = [System.Drawing.Color]::White
    $btnDone.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnDone.SetBounds(20, 98, 100, 34)
    $f.Controls.Add($btnDone)

    $btnSnooze = New-Object System.Windows.Forms.Button
    $btnSnooze.Text = '10 分钟后再提醒'
    $btnSnooze.Font = $script:UiFont
    $btnSnooze.FlatStyle = 'Flat'
    $btnSnooze.FlatAppearance.BorderSize = 1
    $btnSnooze.FlatAppearance.BorderColor = $script:Line
    $btnSnooze.BackColor = [System.Drawing.Color]::White
    $btnSnooze.ForeColor = $script:Ink
    $btnSnooze.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnSnooze.SetBounds(128, 98, 156, 34)
    $f.Controls.Add($btnSnooze)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = '×'
    $btnClose.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 13)
    $btnClose.FlatStyle = 'Flat'
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.BackColor = [System.Drawing.Color]::White
    $btnClose.ForeColor = $script:Ink2
    $btnClose.SetBounds(348, 6, 32, 28)
    $f.Controls.Add($btnClose)

    $f.Add_Paint({
        param($s, $e)
        $p = New-Object System.Drawing.Pen($script:Line, 1)
        $e.Graphics.DrawRectangle($p, 0, 0, ($s.Width - 1), ($s.Height - 1))
        $p.Dispose()
    })

    $btnDone.Add_Click({
        $fm = $this.FindForm()
        Complete-TaskById -Id ([string]$fm.Tag)
        Update-List
        $fm.Close()
    })

    $btnSnooze.Add_Click({
        $fm = $this.FindForm()
        Snooze-TaskById -Id ([string]$fm.Tag) -Minutes 10
        Update-List
        $fm.Close()
    })

    $btnClose.Add_Click({ $this.FindForm().Close() })

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 25000
    $timer.Tag = $f
    $timer.Add_Tick({
        $tmr = $this
        $tmr.Stop()
        if ($tmr.Tag) { try { $tmr.Tag.Close() } catch { } }
        $tmr.Dispose()
    })
    $timer.Start()

    $f.Add_FormClosed({
        param($s, $e)
        if ($script:ToastForms.Contains($s)) { [void]$script:ToastForms.Remove($s) }
        Reposition-Toasts
    })

    [void]$script:ToastForms.Add($f)
    $f.Show()
    Reposition-Toasts

    if ($script:Settings.Sound) {
        try { [System.Media.SystemSounds]::Asterisk.Play() } catch { }
    }
}

# ================================================================ 主界面

function Update-Stats {
    if (-not $script:Form) { return }
    $now = Get-Date
    $todo = 0; $over = 0; $done = 0
    foreach ($t in $script:Tasks) {
        if ([bool]$t.done) { $done++; continue }
        $todo++
        $e = Get-TaskDeadline $t
        if ($e -and $e -lt $now) { $over++ }
    }
    $s = "待办 $todo 项"
    if ($over -gt 0) { $s = $s + " · 逾期 $over 项" }
    if ($done -gt 0) { $s = $s + " · 已完成 $done 项" }
    $script:LblStats.Text = $s
    if ($over -gt 0) { $script:LblStats.ForeColor = $script:Danger } else { $script:LblStats.ForeColor = $script:Ink2 }
}

function Update-List {
    if (-not $script:ListView) { return }
    $script:Updating = $true
    $lv = $script:ListView
    $sel = $null
    if ($lv.SelectedItems.Count -gt 0) { $sel = [string]$lv.SelectedItems[0].Tag }

    $lv.BeginUpdate()
    $lv.Items.Clear()

    $visible = @($script:Tasks | Where-Object {
            if ($script:Filter -eq 'todo') { -not [bool]$_.done }
            elseif ($script:Filter -eq 'done') { [bool]$_.done }
            else { $true }
        } | Sort-Object -Property @{ Expression = { if ([bool]$_.done) { 1 } else { 0 } } }, @{ Expression = { $d = ConvertTo-DateTime $_.due; if ($d) { $d } else { [datetime]::MaxValue } } })

    foreach ($t in $visible) {
        $due = ConvertTo-DateTime $t.due
        $li = New-Object System.Windows.Forms.ListViewItem([string]$t.title)
        if ($due) { [void]$li.SubItems.Add((Get-DueText -Due $due -Done ([bool]$t.done) -End (ConvertTo-DateTime $t.end))) }
        else { [void]$li.SubItems.Add('-') }
        [void]$li.SubItems.Add((Get-RepeatLabel ([string]$t.repeat)))
        $li.Tag = [string]$t.id
        $li.Checked = [bool]$t.done

        $deadline = Get-TaskDeadline $t
        if ([bool]$t.done) {
            $li.ForeColor = [System.Drawing.Color]::FromArgb(152, 160, 172)
            $li.Font = $script:StrikeFont
        } elseif ($deadline -and $deadline -lt (Get-Date)) {
            $li.ForeColor = $script:Danger
        } elseif ($due -and $t.end -and $due -le (Get-Date)) {
            $li.ForeColor = $script:Accent
        } else {
            $li.ForeColor = $script:Ink
        }
        [void]$lv.Items.Add($li)
    }
    $lv.EndUpdate()

    if ($sel) {
        foreach ($i in $lv.Items) { if ([string]$i.Tag -eq $sel) { $i.Selected = $true; break } }
    }

    $script:EmptyLabel.Visible = ($lv.Items.Count -eq 0)
    if ($lv.Items.Count -eq 0) {
        if ($script:Filter -eq 'done') { $script:EmptyLabel.Text = "还没有已完成的任务`r`n勾选左侧方框即可完成" }
        else { $script:EmptyLabel.Text = "今天还没有任务`r`n在上方输入框写下要做的事，回车即可添加" }
    }
    $script:Updating = $false
    Update-Stats
}

function Set-Filter {
    param([string]$Name)
    $script:Filter = $Name
    foreach ($k in @('all', 'todo', 'done')) {
        $b = $script:TabButtons[$k]
        if ($k -eq $Name) { $b.ForeColor = $script:Accent; $b.Font = $script:UiBold }
        else { $b.ForeColor = $script:Ink2; $b.Font = $script:UiFont }
    }
    Update-List
}

function Add-TaskFromInput {
    $title = $script:TxtTitle.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($title)) {
        $script:TxtTitle.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 238)
        $script:TxtTitle.Focus()
        return
    }
    $script:TxtTitle.BackColor = [System.Drawing.Color]::White

    # 用户没动过时间控件时，用「此刻」重新算默认值，避免常驻一天后默认值停在启动时刻
    $script:AddSlotSyncing = $true
    try {
        $def = Get-AddSlot
        if (-not $script:AddTouchedDate) { $script:DtpDate.Value = $def.Date }
        if (-not $script:AddTouchedStart) { $script:DtpStart.Value = $def }
        if (-not $script:AddTouchedEnd) { $script:DtpEnd.Value = $def.AddHours(1) }
    } finally {
        $script:AddSlotSyncing = $false
    }

    $date = $script:DtpDate.Value.Date
    $st = $script:DtpStart.Value
    $due = $date.AddHours($st.Hour).AddMinutes($st.Minute)

    $endText = $null
    if ($script:ChkRange.Checked) {
        $et = $script:DtpEnd.Value
        $endDt = $date.AddHours($et.Hour).AddMinutes($et.Minute)
        if ($endDt -le $due) { $endDt = $endDt.AddDays(1) }
        $endText = $endDt.ToString('yyyy-MM-ddTHH:mm:ss')
    }

    $repeat = 'none'
    if ($script:CbRepeat.SelectedIndex -gt 0) { $repeat = $script:RepeatKeys[$script:CbRepeat.SelectedIndex] }

    $task = [pscustomobject]@{
        id = [guid]::NewGuid().ToString('N'); title = $title; note = ''
        due = $due.ToString('yyyy-MM-ddTHH:mm:ss'); end = $endText; repeat = $repeat; lead = 0
        done = $false; doneAt = $null; notified = $false; snoozeUntil = $null
        lastAlert = $null
        createdAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'); completedCount = 0
    }
    [void]$script:Tasks.Add($task)

    if (-not (Save-Tasks)) {
        [void]$script:Tasks.Remove($task)
        $script:TxtTitle.Focus()
        return
    }

    $script:TxtTitle.Text = ''
    # 没动过的字段滚到下一个默认时段；手选过的保留，方便连续添加同一时间
    $script:AddSlotSyncing = $true
    try {
        $nextSlot = Get-AddSlot
        if (-not $script:AddTouchedDate) { $script:DtpDate.Value = $nextSlot.Date }
        if (-not $script:AddTouchedStart) { $script:DtpStart.Value = $nextSlot }
        if (-not $script:AddTouchedEnd) { $script:DtpEnd.Value = $nextSlot.AddHours(1) }
    } finally {
        $script:AddSlotSyncing = $false
    }

    $label = $due.ToString('HH:mm')
    if ($endText) {
        $ed = ConvertTo-DateTime $endText
        if ($ed) { $label = $label + '-' + $ed.ToString('HH:mm') }
    }
    $script:Flash = "已添加 $label"

    if ($script:Filter -eq 'done') { Set-Filter 'todo' } else { Update-List }
    $script:TxtTitle.Focus()
}

function Invoke-PendingCheck {
    $p = $script:PendingCheck
    $script:PendingCheck = $null
    if (-not $p) { return }
    $t = Get-TaskById $p.Id
    if (-not $t) { return }
    if ($p.Checked) {
        Complete-TaskById -Id $p.Id
    } else {
        $t.done = $false
        $t.doneAt = $null
        $t.notified = $false
        $t.snoozeUntil = $null
        $t | Add-Member -NotePropertyName lastAlert -NotePropertyValue $null -Force
        Save-Tasks
        $script:Flash = '已恢复为待办'
    }
    Update-List
}

function Show-Editor {
    param([string]$TaskId)

    $task = $null
    if ($TaskId) { $task = Get-TaskById $TaskId }

    $dlg = New-Object System.Windows.Forms.Form
    if ($task) { $dlg.Text = '编辑任务' } else { $dlg.Text = '新建任务' }
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.ClientSize = New-Object System.Drawing.Size(420, 366)
    $dlg.BackColor = [System.Drawing.Color]::White
    $dlg.Font = $script:UiFont
    $dlg.ShowInTaskbar = $false

    $mkLabel = {
        param($text, $x, $y, $w)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.ForeColor = $script:Ink2
        $l.AutoSize = $false; $l.BackColor = [System.Drawing.Color]::Transparent
        $l.SetBounds($x, $y, $w, 18)
        $dlg.Controls.Add($l)
    }

    & $mkLabel '任务内容' 20 16 200
    $txtT = New-Object System.Windows.Forms.TextBox
    $txtT.SetBounds(20, 36, 380, 28)
    $txtT.BorderStyle = 'FixedSingle'
    $txtT.Font = $script:UiFont
    $dlg.Controls.Add($txtT)

    & $mkLabel '备注（可选）' 20 74 200
    $txtN = New-Object System.Windows.Forms.TextBox
    $txtN.SetBounds(20, 94, 380, 28)
    $txtN.BorderStyle = 'FixedSingle'
    $txtN.Font = $script:UiFont
    $dlg.Controls.Add($txtN)

    & $mkLabel '日期与时间' 20 132 200
    $dtp = New-Object System.Windows.Forms.DateTimePicker
    $dtp.Format = 'Custom'
    $dtp.CustomFormat = 'yyyy年M月d日 ddd'
    $dtp.SetBounds(20, 152, 168, 26)
    $dlg.Controls.Add($dtp)

    $tStart = New-Object System.Windows.Forms.DateTimePicker
    $tStart.Format = 'Custom'
    $tStart.CustomFormat = 'HH:mm'
    $tStart.ShowUpDown = $true
    $tStart.SetBounds(196, 152, 82, 26)
    $dlg.Controls.Add($tStart)

    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = '至'
    $lblTo.ForeColor = $script:Ink2
    $lblTo.AutoSize = $false
    $lblTo.TextAlign = 'MiddleCenter'
    $lblTo.SetBounds(284, 154, 16, 22)
    $dlg.Controls.Add($lblTo)

    $tEnd = New-Object System.Windows.Forms.DateTimePicker
    $tEnd.Format = 'Custom'
    $tEnd.CustomFormat = 'HH:mm'
    $tEnd.ShowUpDown = $true
    $tEnd.Enabled = $false
    $tEnd.SetBounds(306, 152, 82, 26)
    $dlg.Controls.Add($tEnd)

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = '时间段'
    $chk.ForeColor = $script:Ink2
    $chk.SetBounds(20, 188, 100, 24)
    $dlg.Controls.Add($chk)

    $lblRange = New-Object System.Windows.Forms.Label
    $lblRange.Text = '勾选后任务变成一段（开始 → 结束）'
    $lblRange.ForeColor = [System.Drawing.Color]::FromArgb(150, 158, 170)
    $lblRange.AutoSize = $false
    $lblRange.SetBounds(116, 191, 284, 20)
    $dlg.Controls.Add($lblRange)

    & $mkLabel '重复' 20 220 120
    $cbR = New-Object System.Windows.Forms.ComboBox
    $cbR.DropDownStyle = 'DropDownList'
    $cbR.SetBounds(20, 240, 170, 26)
    $repeatKeys = @('none', 'daily', 'weekly', 'monthly', 'weekday')
    $repeatTexts = @('不重复', '每天', '每周', '每月', '工作日')
    foreach ($x in $repeatTexts) { [void]$cbR.Items.Add($x) }
    $dlg.Controls.Add($cbR)

    & $mkLabel '提前提醒' 206 220 160
    $cbL = New-Object System.Windows.Forms.ComboBox
    $cbL.DropDownStyle = 'DropDownList'
    $cbL.SetBounds(206, 240, 176, 26)
    $leads = @(0, 5, 10, 15, 30, 60, 120, 1440)
    $leadTexts = @('准时', '提前 5 分钟', '提前 10 分钟', '提前 15 分钟', '提前 30 分钟', '提前 1 小时', '提前 2 小时', '提前 1 天')
    foreach ($x in $leadTexts) { [void]$cbL.Items.Add($x) }
    $dlg.Controls.Add($cbL)

    $lblErr = New-Object System.Windows.Forms.Label
    $lblErr.SetBounds(20, 278, 380, 18)
    $lblErr.ForeColor = $script:Danger
    $lblErr.Text = ''
    $dlg.Controls.Add($lblErr)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = '保存'
    $btnOk.FlatStyle = 'Flat'
    $btnOk.FlatAppearance.BorderSize = 0
    $btnOk.BackColor = $script:Accent
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.Font = $script:UiBold
    $btnOk.SetBounds(214, 314, 86, 34)
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = '取消'
    $btnCancel.FlatStyle = 'Flat'
    $btnCancel.FlatAppearance.BorderSize = 1
    $btnCancel.FlatAppearance.BorderColor = $script:Line
    $btnCancel.BackColor = [System.Drawing.Color]::White
    $btnCancel.SetBounds(310, 314, 86, 34)
    $btnCancel.DialogResult = 'Cancel'
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    $baseTime = (Get-Date)
    if ($task) {
        $txtT.Text = [string]$task.title
        $txtN.Text = [string]$task.note
        $d = ConvertTo-DateTime $task.due
        if ($d) {
            $dtp.Value = $d.Date
            $tStart.Value = $d.Date.AddHours($d.Hour).AddMinutes($d.Minute)
            $baseTime = $d
        } else {
            $dtp.Value = (Get-Date)
            $tStart.Value = (Get-Date)
        }
        $ed = ConvertTo-DateTime $task.end
        if ($ed) {
            $chk.Checked = $true
            $tEnd.Enabled = $true
            $tEnd.Value = $ed.Date.AddHours($ed.Hour).AddMinutes($ed.Minute)
        } else {
            $tEnd.Value = $baseTime.AddHours(1)
        }
        $ri = $repeatKeys.IndexOf([string]$task.repeat)
        if ($ri -lt 0) { $ri = 0 }
        $cbR.SelectedIndex = $ri
        $li = $leads.IndexOf([int]$task.lead)
        if ($li -lt 0) { $li = 0 }
        $cbL.SelectedIndex = $li
    } else {
        $dtp.Value = (Get-Date)
        $n = (Get-Date)
        if ($n.Minute -lt 30) { $n = $n.Date.AddHours($n.Hour).AddMinutes(30) } else { $n = $n.Date.AddHours($n.Hour).AddMinutes(60) }
        $tStart.Value = $n
        $tEnd.Value = $n.AddHours(1)
        $cbR.SelectedIndex = 0
        $cbL.SelectedIndex = 0
    }

    $chk.Add_CheckedChanged({ $this.FindForm().Tag.End.Enabled = $this.Checked })

    $ctx = @{
        Txt = $txtT; Note = $txtN; Date = $dtp; Start = $tStart; End = $tEnd; Range = $chk
        Repeat = $cbR; Lead = $cbL; Err = $lblErr
        Keys = $repeatKeys; Leads = $leads; TaskId = $TaskId
    }
    $dlg.Tag = $ctx

    $btnOk.Add_Click({
        $c = $this.FindForm().Tag
        $title = $c.Txt.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($title)) { $c.Err.Text = '请填写任务内容'; return }

        $date = $c.Date.Value.Date
        $st = $c.Start.Value
        $due = $date.AddHours($st.Hour).AddMinutes($st.Minute)

        $endText = $null
        if ($c.Range.Checked) {
            $et = $c.End.Value
            $endDt = $date.AddHours($et.Hour).AddMinutes($et.Minute)
            if ($endDt -le $due) { $endDt = $endDt.AddDays(1) }
            $endText = $endDt.ToString('yyyy-MM-ddTHH:mm:ss')
        }

        $repeat = [string]$c.Keys[$c.Repeat.SelectedIndex]
        $lead = [int]$c.Leads[$c.Lead.SelectedIndex]

        $t = $null
        if ($c.TaskId) { $t = Get-TaskById ([string]$c.TaskId) }
        if ($t) {
            $t.title = $title
            $t.note = $c.Note.Text.Trim()
            $t.due = $due.ToString('yyyy-MM-ddTHH:mm:ss')
            $t | Add-Member -NotePropertyName end -NotePropertyValue $endText -Force
            $t.repeat = $repeat
            $t.lead = $lead
            if ([bool]$t.done) { $t.done = $false; $t.doneAt = $null }
            $t.notified = $false
            $t.snoozeUntil = $null
            $t | Add-Member -NotePropertyName lastAlert -NotePropertyValue $null -Force
            $script:Flash = '已保存修改'
        } else {
            $obj = [pscustomobject]@{
                id = [guid]::NewGuid().ToString('N'); title = $title; note = $c.Note.Text.Trim()
                due = $due.ToString('yyyy-MM-ddTHH:mm:ss'); end = $endText; repeat = $repeat; lead = $lead
                done = $false; doneAt = $null; notified = $false; snoozeUntil = $null
                createdAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'); completedCount = 0
            }
            [void]$script:Tasks.Add($obj)
            $script:Flash = "已添加：$title"
        }
        Save-Tasks
        $this.FindForm().DialogResult = 'OK'
        $this.FindForm().Close()
    })

    $r = $dlg.ShowDialog($script:Form)
    $dlg.Dispose()
    if ($r -eq 'OK') { Update-List }
}

function Show-Settings {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = '设置'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.ClientSize = New-Object System.Drawing.Size(424, 330)
    $dlg.BackColor = [System.Drawing.Color]::White
    $dlg.Font = $script:UiFont
    $dlg.ShowInTaskbar = $false

    $chkAuto = New-Object System.Windows.Forms.CheckBox
    $chkAuto.Text = '开机自动启动（登录 Windows 后静默运行在托盘）'
    $chkAuto.SetBounds(20, 18, 380, 26)
    $chkAuto.Checked = Test-Autostart
    $dlg.Controls.Add($chkAuto)

    $chkSound = New-Object System.Windows.Forms.CheckBox
    $chkSound.Text = '提醒时播放提示音'
    $chkSound.SetBounds(20, 52, 380, 26)
    $chkSound.Checked = [bool]$script:Settings.Sound
    $dlg.Controls.Add($chkSound)

    $chkTray = New-Object System.Windows.Forms.CheckBox
    $chkTray.Text = '关闭窗口时最小化到托盘（不退出）'
    $chkTray.SetBounds(20, 86, 380, 26)
    $chkTray.Checked = [bool]$script:Settings.MinimizeToTray
    $dlg.Controls.Add($chkTray)

    $lblNag = New-Object System.Windows.Forms.Label
    $lblNag.Text = '逾期后每隔'
    $lblNag.ForeColor = $script:Ink2
    $lblNag.AutoSize = $false
    $lblNag.SetBounds(20, 122, 80, 22)
    $dlg.Controls.Add($lblNag)

    $nagValues = @(0, 10, 15, 30, 60)
    $nagTexts = @('不催','10 分钟','15 分钟','30 分钟','60 分钟')
    $cbNag = New-Object System.Windows.Forms.ComboBox
    $cbNag.DropDownStyle = 'DropDownList'
    $cbNag.SetBounds(100, 119, 96, 26)
    foreach ($x in $nagTexts) { [void]$cbNag.Items.Add($x) }
    $cur = [int]$script:Settings.NagMinutes
    $idx = $nagValues.IndexOf($cur)
    if ($idx -lt 0) { $idx = 2 }
    $cbNag.SelectedIndex = $idx
    $dlg.Controls.Add($cbNag)

    $lblNag2 = New-Object System.Windows.Forms.Label
    $lblNag2.Text = '再提醒一次（任务没完成才会催）'
    $lblNag2.ForeColor = [System.Drawing.Color]::FromArgb(150, 158, 170)
    $lblNag2.AutoSize = $false
    $lblNag2.SetBounds(202, 122, 202, 22)
    $dlg.Controls.Add($lblNag2)

    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text = "数据文件：$($script:DataFile)"
    $lblPath.ForeColor = $script:Ink2
    $lblPath.AutoSize = $false
    $lblPath.SetBounds(20, 156, 380, 36)
    $dlg.Controls.Add($lblPath)

    $btnFolder = New-Object System.Windows.Forms.Button
    $btnFolder.Text = '打开数据文件夹'
    $btnFolder.FlatStyle = 'Flat'
    $btnFolder.FlatAppearance.BorderSize = 1
    $btnFolder.FlatAppearance.BorderColor = $script:Line
    $btnFolder.SetBounds(20, 196, 130, 32)
    $btnFolder.Add_Click({ Start-Process explorer.exe -ArgumentList $script:DataDir })
    $dlg.Controls.Add($btnFolder)

    $lblVer = New-Object System.Windows.Forms.Label
    $lblVer.Text = "$($script:AppName) v$($script:Version) · 纯本地运行，不联网"
    $lblVer.ForeColor = $script:Ink2
    $lblVer.SetBounds(20, 240, 380, 18)
    $dlg.Controls.Add($lblVer)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = '保存'
    $btnSave.FlatStyle = 'Flat'
    $btnSave.FlatAppearance.BorderSize = 0
    $btnSave.BackColor = $script:Accent
    $btnSave.ForeColor = [System.Drawing.Color]::White
    $btnSave.Font = $script:UiBold
    $btnSave.SetBounds(218, 272, 86, 32)
    $dlg.Controls.Add($btnSave)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = '关闭'
    $btnClose.FlatStyle = 'Flat'
    $btnClose.FlatAppearance.BorderSize = 1
    $btnClose.FlatAppearance.BorderColor = $script:Line
    $btnClose.DialogResult = 'Cancel'
    $btnClose.SetBounds(314, 272, 86, 32)
    $dlg.Controls.Add($btnClose)
    $dlg.CancelButton = $btnClose

    $dlg.Tag = @{ Auto = $chkAuto; Sound = $chkSound; Tray = $chkTray; Nag = $cbNag; NagValues = $nagValues }

    $btnSave.Add_Click({
        $c = $this.FindForm().Tag
        $script:Settings.Sound = [bool]$c.Sound.Checked
        $script:Settings.MinimizeToTray = [bool]$c.Tray.Checked
        $script:Settings.NagMinutes = [int]$c.NagValues[$c.Nag.SelectedIndex]
        Save-Settings
        try {
            $on = Set-Autostart -Enabled ([bool]$c.Auto.Checked)
            if ($on) { $script:Flash = '已开启开机自启' } else { $script:Flash = '已关闭开机自启' }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("写入启动文件夹失败：`r`n$($_.Exception.Message)", $script:AppName)
        }
        $this.FindForm().Close()
    })

    [void]$dlg.ShowDialog($script:Form)
    $dlg.Dispose()
}

function Show-MainWindow {
    if (-not $script:Form) { return }
    if ($script:Form.WindowState -eq 'Minimized') { $script:Form.WindowState = 'Normal' }
    $script:Form.Show()
    $script:Form.Activate()
    $script:Form.BringToFront()
}

function Build-Ui {
    $script:StrikeFont = New-Object System.Drawing.Font($script:UiFont, [System.Drawing.FontStyle]::Strikeout)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $script:AppName
    $form.StartPosition = 'CenterScreen'
    $form.ClientSize = New-Object System.Drawing.Size(484, 664)
    $form.MinimumSize = New-Object System.Drawing.Size(484, 600)
    $form.BackColor = $script:Bg
    $form.Font = $script:UiFont
    $form.KeyPreview = $true
    if (Test-Path $script:IconPath) { try { $form.Icon = New-Object System.Drawing.Icon($script:IconPath) } catch { } }
    $script:Form = $form

    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'
    $root.ColumnCount = 1
    $root.RowCount = 5
    $root.Margin = New-Object System.Windows.Forms.Padding(0)
    $root.Padding = New-Object System.Windows.Forms.Padding(0)
    $root.BackColor = $script:Bg
    [void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 74)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 126)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
    $form.Controls.Add($root)

    # ---------- 标题
    $head = New-Object System.Windows.Forms.Panel
    $head.Dock = 'Fill'
    $head.Margin = New-Object System.Windows.Forms.Padding(0)
    $head.BackColor = $script:Bg
    $root.Controls.Add($head, 0, 0)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = $script:AppName
    $lblTitle.Font = $script:BigFont
    $lblTitle.ForeColor = $script:Ink
    $lblTitle.AutoSize = $false
    $lblTitle.BackColor = [System.Drawing.Color]::Transparent
    $lblTitle.SetBounds(22, 14, 240, 32)
    $head.Controls.Add($lblTitle)

    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text = (Get-Date).ToString('M月d日 dddd')
    $lblSub.ForeColor = $script:Ink2
    $lblSub.AutoSize = $false
    $lblSub.BackColor = [System.Drawing.Color]::Transparent
    $lblSub.SetBounds(24, 46, 240, 18)
    $head.Controls.Add($lblSub)

    $btnSet = New-Object System.Windows.Forms.Button
    $btnSet.Text = '设置'
    $btnSet.FlatStyle = 'Flat'
    $btnSet.FlatAppearance.BorderSize = 1
    $btnSet.FlatAppearance.BorderColor = $script:Line
    $btnSet.BackColor = [System.Drawing.Color]::White
    $btnSet.ForeColor = $script:Ink2
    $btnSet.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnSet.SetBounds(412, 22, 50, 28)
    $btnSet.Add_Click({ Show-Settings })
    $head.Controls.Add($btnSet)

    $headLine = New-Object System.Windows.Forms.Panel
    $headLine.Dock = 'Bottom'
    $headLine.Height = 1
    $headLine.BackColor = $script:Line
    $head.Controls.Add($headLine)

    # ---------- 添加区
    $add = New-Object System.Windows.Forms.Panel
    $add.Dock = 'Fill'
    $add.Margin = New-Object System.Windows.Forms.Padding(0)
    $add.BackColor = $script:SoftBg
    $root.Controls.Add($add, 0, 1)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.BorderStyle = 'FixedSingle'
    $txt.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10.5)
    $txt.BackColor = [System.Drawing.Color]::White
    $txt.SetBounds(14, 12, 456, 30)
    $add.Controls.Add($txt)

    $dtp = New-Object System.Windows.Forms.DateTimePicker
    $dtp.Format = 'Custom'
    $dtp.CustomFormat = 'M月d日 ddd'
    $dtp.Font = $script:UiFont
    $dtp.SetBounds(14, 50, 112, 26)
    $add.Controls.Add($dtp)

    $tStart = New-Object System.Windows.Forms.DateTimePicker
    $tStart.Format = 'Custom'
    $tStart.CustomFormat = 'HH:mm'
    $tStart.ShowUpDown = $true
    $tStart.Font = $script:UiFont
    $tStart.SetBounds(132, 50, 72, 26)
    $slot = Get-AddSlot
    $dtp.Value = $slot.Date
    $tStart.Value = $slot
    $add.Controls.Add($tStart)

    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = '至'
    $lblTo.ForeColor = $script:Ink2
    $lblTo.AutoSize = $false
    $lblTo.TextAlign = 'MiddleCenter'
    $lblTo.SetBounds(208, 52, 16, 22)
    $add.Controls.Add($lblTo)

    $tEnd = New-Object System.Windows.Forms.DateTimePicker
    $tEnd.Format = 'Custom'
    $tEnd.CustomFormat = 'HH:mm'
    $tEnd.ShowUpDown = $true
    $tEnd.Font = $script:UiFont
    $tEnd.Enabled = $false
    $tEnd.SetBounds(228, 50, 72, 26)
    $tEnd.Value = $slot.AddHours(1)
    $add.Controls.Add($tEnd)

    $script:AddTouchedDate = $false
    $script:AddTouchedStart = $false
    $script:AddTouchedEnd = $false
    $script:AddSlotSyncing = $false
    $dtp.Add_ValueChanged({ if (-not $script:AddSlotSyncing) { $script:AddTouchedDate = $true } })
    $tStart.Add_ValueChanged({ if (-not $script:AddSlotSyncing) { $script:AddTouchedStart = $true } })
    $tEnd.Add_ValueChanged({ if (-not $script:AddSlotSyncing) { $script:AddTouchedEnd = $true } })

    $cbR = New-Object System.Windows.Forms.ComboBox
    $cbR.DropDownStyle = 'DropDownList'
    $cbR.Font = $script:UiFont
    $cbR.SetBounds(306, 50, 84, 26)
    foreach ($x in $script:RepeatLabels) { [void]$cbR.Items.Add($x) }
    $cbR.SelectedIndex = 0
    $add.Controls.Add($cbR)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = '添加'
    $btnAdd.FlatStyle = 'Flat'
    $btnAdd.FlatAppearance.BorderSize = 0
    $btnAdd.BackColor = $script:Accent
    $btnAdd.ForeColor = [System.Drawing.Color]::White
    $btnAdd.Font = $script:UiBold
    $btnAdd.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnAdd.SetBounds(396, 50, 74, 27)
    $btnAdd.Add_Click({ Add-TaskFromInput })
    $add.Controls.Add($btnAdd)

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = '时间段'
    $chk.ForeColor = $script:Ink2
    $chk.Font = $script:UiFont
    $chk.SetBounds(14, 86, 86, 24)
    $chk.Add_CheckedChanged({ $script:DtpEnd.Enabled = $this.Checked })
    $add.Controls.Add($chk)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = '不勾选只提醒开始时间；勾选后记录「开始 → 结束」'
    $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(150, 158, 170)
    $lblHint.AutoSize = $false
    $lblHint.SetBounds(104, 89, 366, 20)
    $add.Controls.Add($lblHint)

    # ---------- 筛选
    $tabBar = New-Object System.Windows.Forms.Panel
    $tabBar.Dock = 'Fill'
    $tabBar.Margin = New-Object System.Windows.Forms.Padding(0)
    $tabBar.BackColor = $script:Bg
    $root.Controls.Add($tabBar, 0, 2)

    $script:TabButtons = @{}
    $tabDefs = @(
        @('all', '全部', 16),
        @('todo', '待办', 76),
        @('done', '已完成', 136)
    )
    foreach ($d in $tabDefs) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $d[1]
        $b.FlatStyle = 'Flat'
        $b.FlatAppearance.BorderSize = 0
        $b.BackColor = $script:Bg
        $b.Cursor = [System.Windows.Forms.Cursors]::Hand
        $b.SetBounds($d[2], 9, 56, 26)
        $b.Tag = $d[0]
        $b.Add_Click({ Set-Filter ([string]$this.Tag) })
        $tabBar.Controls.Add($b)
        $script:TabButtons[$d[0]] = $b
    }

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = '清除已完成'
    $btnClear.FlatStyle = 'Flat'
    $btnClear.FlatAppearance.BorderSize = 0
    $btnClear.BackColor = $script:Bg
    $btnClear.ForeColor = $script:Ink2
    $btnClear.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClear.SetBounds(362, 9, 106, 26)
    $btnClear.Add_Click({
        $removed = 0
        for ($i = $script:Tasks.Count - 1; $i -ge 0; $i--) {
            if ([bool]$script:Tasks[$i].done) { $script:Tasks.RemoveAt($i); $removed++ }
        }
        if ($removed -gt 0) {
            Save-Tasks
            $script:Flash = "已清除 $removed 条已完成任务"
            Update-List
        }
    })
    $tabBar.Controls.Add($btnClear)

    $tabLine = New-Object System.Windows.Forms.Panel
    $tabLine.Dock = 'Bottom'
    $tabLine.Height = 1
    $tabLine.BackColor = $script:Line
    $tabBar.Controls.Add($tabLine)

    # ---------- 列表
    $listWrap = New-Object System.Windows.Forms.Panel
    $listWrap.Dock = 'Fill'
    $listWrap.Margin = New-Object System.Windows.Forms.Padding(0)
    $listWrap.BackColor = $script:Bg
    $listWrap.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
    $root.Controls.Add($listWrap, 0, 3)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Dock = 'Fill'
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.CheckBoxes = $true
    $lv.MultiSelect = $false
    $lv.BorderStyle = 'None'
    $lv.HideSelection = $false
    $lv.HeaderStyle = 'Nonclickable'
    $lv.Font = $script:UiFont
    $lv.BackColor = [System.Drawing.Color]::White
    $lv.ForeColor = $script:Ink
    [void]$lv.Columns.Add('任务', 194)
    [void]$lv.Columns.Add('时间', 178)
    [void]$lv.Columns.Add('重复', 60)
    $il = New-Object System.Windows.Forms.ImageList
    $il.ImageSize = New-Object System.Drawing.Size(1, 26)
    $lv.SmallImageList = $il
    $listWrap.Controls.Add($lv)

    $empty = New-Object System.Windows.Forms.Label
    $empty.TextAlign = 'MiddleCenter'
    $empty.ForeColor = [System.Drawing.Color]::FromArgb(160, 168, 180)
    $empty.BackColor = [System.Drawing.Color]::White
    $empty.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $empty.Visible = $false
    $empty.Cursor = [System.Windows.Forms.Cursors]::Hand
    $empty.SetBounds(10, 46, 444, 110)
    $empty.Add_Click({ $script:TxtTitle.Focus() })
    $listWrap.Controls.Add($empty)

    # ---------- 底部
    $foot = New-Object System.Windows.Forms.Panel
    $foot.Dock = 'Fill'
    $foot.Margin = New-Object System.Windows.Forms.Padding(0)
    $foot.BackColor = $script:Bg
    $root.Controls.Add($foot, 0, 4)

    $footLine = New-Object System.Windows.Forms.Panel
    $footLine.Dock = 'Top'
    $footLine.Height = 1
    $footLine.BackColor = $script:Line
    $foot.Controls.Add($footLine)

    $lblStats = New-Object System.Windows.Forms.Label
    $lblStats.Text = ''
    $lblStats.ForeColor = $script:Ink2
    $lblStats.AutoSize = $false
    $lblStats.SetBounds(16, 13, 260, 20)
    $foot.Controls.Add($lblStats)

    $lblFlash = New-Object System.Windows.Forms.Label
    $lblFlash.Text = ''
    $lblFlash.ForeColor = $script:Accent
    $lblFlash.Font = $script:UiBold
    $lblFlash.AutoSize = $false
    $lblFlash.SetBounds(196, 13, 274, 20)
    $lblFlash.TextAlign = 'MiddleRight'
    $lblFlash.BackColor = $script:Bg
    $lblFlash.Visible = $false
    $foot.Controls.Add($lblFlash)

    $script:TxtTitle = $txt
    $script:DtpDate = $dtp
    $script:DtpStart = $tStart
    $script:DtpEnd = $tEnd
    $script:ChkRange = $chk
    $script:CbRepeat = $cbR
    $script:ListView = $lv
    $script:LblStats = $lblStats
    $script:LblFlash = $lblFlash
    $script:EmptyLabel = $empty

    # ---------- 事件
    $txt.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq 'Return') { $e.SuppressKeyPress = $true; Add-TaskFromInput }
    })
    $txt.Add_TextChanged({
        if ($script:TxtTitle.BackColor -ne [System.Drawing.Color]::White) { $script:TxtTitle.BackColor = [System.Drawing.Color]::White }
    })

    $script:AllowCheck = $false
    $lv.Add_MouseDown({
        param($s, $e)
        $script:AllowCheck = $false
        $it = $s.GetItemAt($e.X, $e.Y)
        if ($null -eq $it) { return }
        $lb = $it.GetBounds([System.Windows.Forms.ItemBoundsPortion]::Label)
        if ($e.X -lt $lb.Left) { $script:AllowCheck = $true }
    })

    $lv.Add_ItemCheck({
        param($s, $e)
        if ($script:Updating) { return }
        if (-not $script:AllowCheck) { $e.NewValue = $e.CurrentValue; return }
        $id = [string]$s.Items[$e.Index].Tag
        $want = ([int]$e.NewValue -eq 1)
        $t = Get-TaskById $id
        if (-not $t) { return }
        if ([bool]$t.done -eq $want) { return }
        $script:PendingCheck = @{ Id = $id; Checked = $want }
        try { [void]$script:Form.BeginInvoke([Action]{ Invoke-PendingCheck }) } catch { }
    })

    $lv.Add_MouseDoubleClick({
        param($s, $e)
        Start-Sleep -Milliseconds 1
        if ($s.SelectedItems.Count -eq 0) { return }
        Show-Editor -TaskId ([string]$s.SelectedItems[0].Tag)
    })

    $lv.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Space) { $script:AllowCheck = $true; return }
        if ($e.KeyCode -eq 'Delete' -and $s.SelectedItems.Count -gt 0) {
            $id = [string]$s.SelectedItems[0].Tag
            $t = Get-TaskById $id
            $r = [System.Windows.Forms.MessageBox]::Show("删除任务「$($t.title)」？", $script:AppName, 'YesNo', 'Question')
            if ($r -eq 'Yes') { Remove-TaskById -Id $id; Update-List }
        }
    })

    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $miDone = $menu.Items.Add('完成')
    $miEdit = $menu.Items.Add('编辑…')
    $miSnooze = $menu.Items.Add('10 分钟后再提醒')
    $miSnooze30 = $menu.Items.Add('30 分钟后再提醒')
    [void]$menu.Items.Add('-')
    $miPostpone = $menu.Items.Add('顺延 10 分钟')
    $miTomorrow = $menu.Items.Add('顺延到明天 09:00')
    [void]$menu.Items.Add('-')
    $miDel = $menu.Items.Add('删除')
    $miDone.Add_Click({
        if ($script:ListView.SelectedItems.Count -gt 0) {
            Complete-TaskById -Id ([string]$script:ListView.SelectedItems[0].Tag)
            Update-List
        }
    })
    $miEdit.Add_Click({
        if ($script:ListView.SelectedItems.Count -gt 0) { Show-Editor -TaskId ([string]$script:ListView.SelectedItems[0].Tag) }
    })
    $miSnooze.Add_Click({
        if ($script:ListView.SelectedItems.Count -gt 0) {
            Snooze-TaskById -Id ([string]$script:ListView.SelectedItems[0].Tag) -Minutes 10
            Update-List
        }
    })
    $miSnooze30.Add_Click({
        if ($script:ListView.SelectedItems.Count -gt 0) {
            Snooze-TaskById -Id ([string]$script:ListView.SelectedItems[0].Tag) -Minutes 30
            Update-List
        }
    })
    $miPostpone.Add_Click({
        if ($script:ListView.SelectedItems.Count -gt 0) {
            Move-TaskById -Id ([string]$script:ListView.SelectedItems[0].Tag) -To ((Get-Date).AddMinutes(10))
            $script:Flash = '已顺延 10 分钟'
            Update-List
        }
    })
    $miTomorrow.Add_Click({
        if ($script:ListView.SelectedItems.Count -gt 0) {
            $tomorrow = (Get-Date).Date.AddDays(1).AddHours(9)
            Move-TaskById -Id ([string]$script:ListView.SelectedItems[0].Tag) -To $tomorrow
            $script:Flash = '已顺延到明天 09:00'
            Update-List
        }
    })
    $miDel.Add_Click({
        if ($script:ListView.SelectedItems.Count -gt 0) {
            $id = [string]$script:ListView.SelectedItems[0].Tag
            $t = Get-TaskById $id
            $r = [System.Windows.Forms.MessageBox]::Show("删除任务「$($t.title)」？", $script:AppName, 'YesNo', 'Question')
            if ($r -eq 'Yes') { Remove-TaskById -Id $id; Update-List }
        }
    })
    $lv.ContextMenuStrip = $menu

    # ---------- 托盘
    $tray = New-Object System.Windows.Forms.NotifyIcon
    try { $tray.Icon = New-Object System.Drawing.Icon($script:IconPath) }
    catch { $tray.Icon = [System.Drawing.SystemIcons]::Application }
    $tray.Text = $script:AppName
    $tray.Visible = $true
    $script:Tray = $tray

    $tm = New-Object System.Windows.Forms.ContextMenuStrip
    $tmShow = $tm.Items.Add('打开主界面')
    $tmAdd = $tm.Items.Add('新建任务…')
    [void]$tm.Items.Add('-')
    $script:TmAuto = New-Object System.Windows.Forms.ToolStripMenuItem('开机自动启动')
    $script:TmAuto.Checked = Test-Autostart
    $script:TmAuto.CheckOnClick = $true
    [void]$tm.Items.Add($script:TmAuto)
    [void]$tm.Items.Add('-')
    $tmExit = $tm.Items.Add('退出')

    $tmShow.Add_Click({ Show-MainWindow })
    $tmAdd.Add_Click({ Show-MainWindow; Show-Editor -TaskId '' })
    $tmExit.Add_Click({ $script:Exiting = $true; $script:Form.Close() })
    $script:TmAuto.Add_Click({
        $mi = $script:TmAuto
        $on = $false
        try {
            $on = Set-Autostart -Enabled ([bool]$mi.Checked)
            if ($on) { $script:Flash = '已开启开机自启' } else { $script:Flash = '已关闭开机自启' }
        } catch {
            $mi.Checked = Test-Autostart
            [System.Windows.Forms.MessageBox]::Show("写入启动文件夹失败：`r`n$($_.Exception.Message)", $script:AppName)
            return
        }
        $mi.Checked = $on
    })
    $tray.ContextMenuStrip = $tm
    $tray.Add_MouseDoubleClick({ Show-MainWindow })

    # ---------- 定时检查
    $tick = New-Object System.Windows.Forms.Timer
    $tick.Interval = 10000
    $tick.Add_Tick({
        try {
            $due = Get-DueTasks -Now (Get-Date)
            if ($due.Count -gt 0) {
                foreach ($t in $due) {
                    Mark-Alerted $t
                    if (-not $script:Exiting) { Show-Toast -Task $t }
                }
                Save-Tasks
                Update-List
            } else {
                Update-Stats
            }
            if ($script:Flash) {
                $script:LblFlash.Text = [string]$script:Flash
                $script:LblFlash.Visible = $true
                $script:Flash = $null
                $script:FlashTimer.Stop()
                $script:FlashTimer.Start()
            }
        } catch { }
    })
    $tick.Start()

    # ---------- 关闭行为
    $form.Add_FormClosing({
        param($s, $e)
        if ($e.CloseReason -eq [System.Windows.Forms.CloseReason]::WindowsShutDown) {
            try { $script:Tray.Visible = $false; $script:Tray.Dispose() } catch { }
            return
        }
        if (-not $script:Exiting -and $script:Settings.MinimizeToTray) {
            $e.Cancel = $true
            $s.Hide()
            if (-not $script:TrayTipShown) {
                $script:TrayTipShown = $true
                try { $script:Tray.ShowBalloonTip(3000, $script:AppName, '已最小化到托盘，双击图标可重新打开', 'Info') } catch { }
            }
        } else {
            try { $script:Tray.Visible = $false; $script:Tray.Dispose() } catch { }
        }
    })
}

# ================================================================ 运行模式

function Invoke-Setup {
    if (-not (Test-Path $script:IconPath)) { New-AppIcon -Path $script:IconPath }
    $vbs = Join-Path $script:AppDir 'Start-TaskReminder.vbs'
    if (-not (Test-Path $vbs)) { throw "缺少 $vbs" }
    New-Shortcut -LinkPath (Get-DesktopShortcut) -VbsPath $vbs -WindowStyle 1
    [void](Set-Autostart -Enabled $true)
    Write-Output 'OK'
}

function Invoke-RemoveShortcuts {
    $a = Get-StartupShortcut
    $d = Get-DesktopShortcut
    if (Test-Path $a) { Remove-Item -LiteralPath $a -Force }
    if (Test-Path $d) { Remove-Item -LiteralPath $d -Force }
    Write-Output 'OK'
}

function Invoke-SelfTest {
    $script:SelfFail = 0
    $ok = {
        param($name, $cond)
        if ($cond) { Write-Host "PASS  $name" }
        else { Write-Host "FAIL  $name"; $script:SelfFail++ }
    }

    $tmpDir = Join-Path $env:TEMP ('trtest_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
    $script:DataDir = $tmpDir
    $script:DataFile = Join-Path $tmpDir 'tasks.json'

    $d = Get-Date '2026-10-23 09:00:00'
    & $ok 'daily +1d' ((Get-NextDue -From $d -Repeat 'daily') -eq $d.AddDays(1))
    & $ok 'weekly +7d' ((Get-NextDue -From $d -Repeat 'weekly') -eq $d.AddDays(7))
    & $ok 'monthly +1m' ((Get-NextDue -From $d -Repeat 'monthly') -eq $d.AddMonths(1))
    & $ok 'weekday fri->mon' ((Get-NextDue -From (Get-Date '2026-10-23 09:00:00') -Repeat 'weekday') -eq (Get-Date '2026-10-26 09:00:00'))
    & $ok 'weekday wed->thu' ((Get-NextDue -From (Get-Date '2026-10-21 09:00:00') -Repeat 'weekday') -eq (Get-Date '2026-10-22 09:00:00'))
    & $ok 'repeat label' ((Get-RepeatLabel 'weekday') -eq '工作日')

    $script:Tasks = New-Object System.Collections.ArrayList
    [void]$script:Tasks.Add([pscustomobject]@{
            id = 'a1'; title = '测试任务'; note = ''; due = '2026-10-23T09:00:00'; repeat = 'none'
            lead = 0; done = $false; doneAt = $null; notified = $false; snoozeUntil = $null
            createdAt = '2026-10-23T08:00:00'; completedCount = 0
        })

    & $ok 'due fires' ((Get-DueTasks -Now (Get-Date '2026-10-23 09:00:30')).Count -eq 1)
    & $ok 'not early' ((Get-DueTasks -Now (Get-Date '2026-10-23 08:59:00')).Count -eq 0)
    $script:Tasks[0].lead = 10
    & $ok 'lead 10min' ((Get-DueTasks -Now (Get-Date '2026-10-23 08:51:00')).Count -eq 1)
    $script:Tasks[0].lead = 0
    $script:Tasks[0].snoozeUntil = '2026-10-23T09:20:00'
    & $ok 'snooze suppresses' ((Get-DueTasks -Now (Get-Date '2026-10-23 09:05:00')).Count -eq 0)
    & $ok 'snooze fires' ((Get-DueTasks -Now (Get-Date '2026-10-23 09:21:00')).Count -eq 1)
    $script:Tasks[0].snoozeUntil = $null

    $script:Tasks[0].repeat = 'daily'
    $script:Tasks[0].notified = $true
    Complete-TaskById -Id 'a1'
    $after = Get-TaskById 'a1'

    & $ok 'repeat keeps time-of-day' (((ConvertTo-DateTime $after.due).TimeOfDay) -eq ([TimeSpan]::FromHours(9)))
    & $ok 'repeat within 8 days' ((ConvertTo-DateTime $after.due) -le (Get-Date).Date.AddDays(8))
    & $ok 'repeat stays pending' (-not $after.done)
    & $ok 'repeat counter' ($after.completedCount -eq 1)
    & $ok 'repeat re-arms' (-not $after.notified)

    $script:Tasks[0].repeat = 'none'
    $script:Tasks[0].done = $false
    Complete-TaskById -Id 'a1'
    & $ok 'once completes' ([bool](Get-TaskById 'a1').done)

    $script:Tasks = New-Object System.Collections.ArrayList
    [void]$script:Tasks.Add([pscustomobject]@{
            id = 'x1'; title = '中文 任务 ✓'; note = '备注 n'; due = '2026-11-01T10:30:00'; repeat = 'weekly'; lead = 5
            done = $false; doneAt = $null; notified = $false; snoozeUntil = $null
            createdAt = '2026-10-23T09:00:00'; completedCount = 2
        })
    [void]$script:Tasks.Add([pscustomobject]@{
            id = 'x2'; title = 'second'; note = ''; due = '2026-11-02T08:00:00'; repeat = 'none'; lead = 0
            done = $true; doneAt = '2026-10-23T09:10:00'; notified = $true; snoozeUntil = $null
            createdAt = '2026-10-23T09:00:00'; completedCount = 0
        })
    [void]$script:Tasks.Add([pscustomobject]@{
            id = 'x3'; title = 'third'; note = ''; due = '2026-11-03T21:15:00'; repeat = 'weekday'; lead = 30
            done = $false; doneAt = $null; notified = $false; snoozeUntil = $null
            createdAt = '2026-10-23T09:00:00'; completedCount = 1
        })
    Save-Tasks
    $loaded = Load-Tasks
    & $ok 'json roundtrip count 3' ($loaded.Count -eq 3)
    & $ok 'json utf8 title' ($loaded[0].title -eq '中文 任务 ✓')
    & $ok 'json fields' ($loaded[0].repeat -eq 'weekly' -and [int]$loaded[0].completedCount -eq 2)
    & $ok 'json order/state' ($loaded[1].done -eq $true -and $loaded[2].repeat -eq 'weekday')

    # 时段语义
    & $ok 'slot running shows 进行中' ((Get-DueText -Due ([datetime]::Now.AddMinutes(-30)) -Done $false -End ([datetime]::Now.AddMinutes(30))) -match '进行中')
    & $ok 'slot ended shows 逾期' ((Get-DueText -Due ([datetime]::Now.AddMinutes(-120)) -Done $false -End ([datetime]::Now.AddMinutes(-60))) -match '逾期')

    # 逾期再催的间隔
    $script:Settings.NagMinutes = 15
    $script:Tasks = New-Object System.Collections.ArrayList
    [void]$script:Tasks.Add([pscustomobject]@{
            id = 'n1'; title = 'nag'; note = ''
            due = ([datetime]::Now.AddMinutes(-60)).ToString('yyyy-MM-ddTHH:mm:ss'); end = $null
            repeat = 'none'; lead = 0; done = $false; doneAt = $null
            notified = $true; snoozeUntil = $null
            lastAlert = ([datetime]::Now.AddMinutes(-5)).ToString('yyyy-MM-ddTHH:mm:ss')
            createdAt = ([datetime]::Now).ToString('yyyy-MM-ddTHH:mm:ss'); completedCount = 0
        })
    & $ok 'nag respects interval' ((Get-DueTasks -Now (Get-Date)).Count -eq 0)
    $script:Tasks[0].lastAlert = ([datetime]::Now.AddMinutes(-20)).ToString('yyyy-MM-ddTHH:mm:ss')
    & $ok 'nag fires after interval' ((Get-DueTasks -Now (Get-Date)).Count -eq 1)

    # snooze 不该再被 lead 提前
    $script:Tasks[0].lead = 30
    $script:Tasks[0].notified = $false
    $script:Tasks[0].snoozeUntil = ([datetime]::Now.AddMinutes(10)).ToString('yyyy-MM-ddTHH:mm:ss')
    & $ok 'snooze ignores lead' ((Get-DueTasks -Now (Get-Date)).Count -eq 0)
    $script:Settings.NagMinutes = 15

    # 重复 + 时段：完成后跨度保持
    $script:Tasks = New-Object System.Collections.ArrayList
    $startDt = [datetime]::Now.AddHours(-1)
    [void]$script:Tasks.Add([pscustomobject]@{
            id = 'r9'; title = 'repeat slot'; note = ''
            due = $startDt.ToString('yyyy-MM-ddTHH:mm:ss'); end = $startDt.AddMinutes(60).ToString('yyyy-MM-ddTHH:mm:ss')
            repeat = 'daily'; lead = 0; done = $false; doneAt = $null
            notified = $true; snoozeUntil = $null; lastAlert = $null
            createdAt = ([datetime]::Now).ToString('yyyy-MM-ddTHH:mm:ss'); completedCount = 0
        })
    Complete-TaskById -Id 'r9'
    $r9 = Get-TaskById 'r9'
    & $ok 'repeat slot keeps span' (((ConvertTo-DateTime $r9.end) - (ConvertTo-DateTime $r9.due)).TotalMinutes -eq 60)

    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue

    & $ok 'startup shortcut points to vbs' ((Test-Path (Join-Path $script:AppDir 'Start-TaskReminder.vbs')))

    if ($script:SelfFail -gt 0) {
        Write-Host "SELFTEST FAILED ($($script:SelfFail) failed)"
        return $false
    }
    Write-Host 'SELFTEST PASSED'
    return $true
}

# ================================================================ main

if ($SelfTest) {
    $pass = Invoke-SelfTest
    if ($pass) { exit 0 } else { exit 1 }
}

if ($Setup) {
    Invoke-Setup
    exit 0
}

if ($RemoveShortcuts) {
    Invoke-RemoveShortcuts
    exit 0
}

if ($Preview) {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
    Load-Settings
    $script:Tasks = New-Object System.Collections.ArrayList
    if (-not (Test-Path $script:IconPath)) { try { New-AppIcon -Path $script:IconPath } catch { } }
    Build-Ui
    $fake = [pscustomobject]@{
        id = 'preview'; title = '这是一条提醒，会显示在屏幕右下角'; note = '可以点「完成」或稍后再提醒'
        due = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'); repeat = 'daily'; lead = 0
        done = $false; doneAt = $null; notified = $false; snoozeUntil = $null
        createdAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'); completedCount = 0
    }
    Show-Toast -Task $fake
    $deadline = (Get-Date).AddSeconds(14)
    while ((Get-Date) -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 60
    }
    exit 0
}

$mutex = New-Object System.Threading.Mutex($false, 'Local\TaskReminderApp_9117')
$hasHandle = $false
try { $hasHandle = $mutex.WaitOne(0, $false) } catch { $hasHandle = $false }
if (-not $hasHandle) {
    try {
        Add-Type -Namespace TrNative -Name Win -MemberDefinition @"
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindow(string cls, string name);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
"@
        $h = [TrNative.Win]::FindWindow($null, $script:AppName)
        if ($h -ne [IntPtr]::Zero) {
            [void][TrNative.Win]::ShowWindow($h, 9)
            [void][TrNative.Win]::SetForegroundWindow($h)
        }
    } catch { }
    Write-Log 'another instance already running - exit'
    exit 0
}

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
$script:StartHidden = [bool]$StartHidden
Write-Log 'boot: visual styles ok'

[System.Windows.Forms.Application]::add_ThreadException({
    param($s, $e)
    Write-Log ('UI EXCEPTION: ' + $e.Exception.ToString())
})

try {
    Load-Settings
    Write-Log 'settings loaded'
    $script:Tasks = Load-Tasks
    Write-Log "tasks loaded: $($script:Tasks.Count)"
    if (-not (Test-Path $script:IconPath)) { try { New-AppIcon -Path $script:IconPath } catch { } }
    Write-Log 'icon ok'

    $script:FlashTimer = New-Object System.Windows.Forms.Timer
    $script:FlashTimer.Interval = 3200
    $script:FlashTimer.Add_Tick({ $script:LblFlash.Visible = $false; $this.Stop() })

    Build-Ui
    Write-Log 'ui built'
    Set-Filter 'todo'
    Update-Stats
    Write-Log 'list rendered'

    $script:Form.Add_Shown({
        Write-Log 'shown'
        $script:TxtTitle.Focus()
        $script:Settings.Launched = [int]$script:Settings.Launched + 1
        Save-Settings

        if ($script:Settings.Launched -le 1 -and -not (Test-Autostart)) {
            try { [void](Set-Autostart -Enabled $true) } catch { }
        }

        $due = Get-DueTasks -Now (Get-Date)
        if ($due.Count -gt 0) {
            foreach ($t in $due) { Mark-Alerted $t; Show-Toast -Task $t }
            Save-Tasks
            Update-List
        }
        if ($script:StartHidden) {
            $script:Form.Hide()
            try { $script:Tray.ShowBalloonTip(2500, $script:AppName, '已在后台运行，双击托盘图标可打开', 'Info') } catch { }
            Write-Log 'started hidden'
        }
        Write-Log 'ready'
    })

    [System.Windows.Forms.Application]::Run($script:Form)
    Write-Log 'run returned (form closed)'
} catch {
    Write-Log ('FATAL: ' + $_.Exception.ToString())
    throw
}

try { $mutex.ReleaseMutex() } catch { }
