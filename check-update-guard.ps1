# =============================================================================
#  check-update-guard.ps1 —— 一键诊断「无法安装必要的更新」
# =============================================================================
#  现象：客户端弹出
#          「无法安装必要的更新 / 检测到应用安装目录下存在用户项目目录…」
#        之后更新永远装不上。
#
#  真实原因（从客户端 app.asar 里挖出来的判定逻辑）：
#        客户端只看「主程序 exe 所在目录」的第一层子目录，
#        只要有一个目录名不在它自己的白名单里，就认定安装目录下有用户内容，
#        然后直接中止整个更新流程。
#        白名单只有：bin locales resources tools log logs _ share-target assets
#                    （以及 $ 开头、nsis 开头的目录）
#        也就是说：安装目录里多出任何一个别的文件夹（不管是不是项目），都会被拦。
#
#  本脚本只做检查：不修改、不删除、不移动任何文件。
# =============================================================================
[CmdletBinding()]
param(
  [string]$InstallDir = '',
  [string]$ReportPath = '',
  [switch]$NoPause
)

$ErrorActionPreference = 'Continue'

# 启动器是纯 ASCII 单行 .bat（不设标题），窗口标题在这里补上
try { if ($Host.UI.RawUI) { $Host.UI.RawUI.WindowTitle = 'WorkBuddy - 更新拦截诊断' } } catch {}
try { [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(936) } catch { }
try { $OutputEncoding = [System.Text.Encoding]::GetEncoding(936) } catch { }

# 客户端自己的白名单（原样照抄，改这里等于改判定标准）
$KnownDirs = @('bin', 'locales', 'resources', 'tools', 'log', 'logs', '_', 'share-target', 'assets')

# 支持的客户端：优先读包内 data/editions.json（v2.0 的唯一事实源），读不到才用内置兜底表。
# 这样国内版/国际版的名称、exe、数据目录、进程规则只维护一份。
$Clients = @()
$EdFile = ''
if ($PSScriptRoot) { $EdFile = Join-Path (Join-Path $PSScriptRoot 'data') 'editions.json' }
if ($EdFile -and (Test-Path -LiteralPath $EdFile)) {
  try {
    $cfg = ([System.IO.File]::ReadAllText($EdFile, [System.Text.Encoding]::UTF8)) | ConvertFrom-Json
    foreach ($k in @('cn', 'intl')) {
      $e = $cfg.editions.$k
      if (-not $e) { continue }
      $Clients += [pscustomobject]@{
        Key         = $k
        Name        = [string]$e.displayName
        Exe         = [string](@($e.winExeNames)[0])
        Data        = [string]$e.dataFolder
        Names       = @($e.winInstallDirNames)
        ProcPattern = [string]$e.winProcNameFilter
      }
    }
  } catch { $Clients = @() }
}
if ($Clients.Count -eq 0) {
  $Clients = @(
    [pscustomobject]@{ Key = 'cn';   Name = 'WorkBuddy（国内版）';     Exe = 'WorkBuddy.exe';   Data = '.workbuddy';    Names = @('WorkBuddy', 'CodeBuddy', 'workbuddy', 'codebuddy'); ProcPattern = '^(?i)(WorkBuddy(?!AI)|CodeBuddy)$' },
    [pscustomobject]@{ Key = 'intl'; Name = 'WorkBuddy AI（国际版）'; Exe = 'WorkBuddyAI.exe'; Data = '.workbuddy-ai'; Names = @('WorkBuddyAI', 'workbuddyai', 'WorkBuddy AI', 'workbuddy-ai'); ProcPattern = '^(?i)(WorkBuddyAI|workbuddy-ai)$' }
  )
}

$script:Report = New-Object System.Collections.ArrayList
$script:ParamError = $false
$anyBlocked = $false
$anyHint = $false

function Add-Report([string]$s) { [void]$script:Report.Add($s) }
function Say([string]$s) { Write-Host $s; Add-Report $s }
function SayBlank() { Write-Host ''; Add-Report '' }
function SayHead([string]$s) {
  Write-Host ''
  Write-Host '  ============================================================' -ForegroundColor Cyan
  Write-Host ("   " + $s) -ForegroundColor Cyan
  Write-Host '  ============================================================' -ForegroundColor Cyan
  Add-Report ''
  Add-Report ('  ' + $s)
  Add-Report ('  ' + ('-' * 60))
}

function J2([string]$a, [string]$b) { return (Join-Path -Path $a -ChildPath $b) }

# 是不是"链接目录"（junction / 符号链接）。
# 客户端用 Node 的 readdir(withFileTypes)，链接目录的 isDirectory() 是 false，
# 所以链接目录不会触发拦截。
function Test-IsLink([string]$path) {
  try {
    $attr = [System.IO.File]::GetAttributes($path)
    return (($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
  } catch { return $false }
}

# 一个目录是不是"这个客户端"的安装目录：
# 客户端自己认的是「主程序 exe 所在目录」，所以这里也以 exe 为准，
# 再用 resources\app.asar 确认是完整安装。
function Test-ClientDir([string]$dir, $client) {
  if (-not $dir) { return $false }
  if (-not (Test-Path -LiteralPath $dir)) { return $false }
  if (-not (Test-Path -LiteralPath (J2 (J2 $dir 'resources') 'app.asar'))) { return $false }
  if (-not (Test-Path -LiteralPath (J2 $dir $client.Exe))) { return $false }
  return $true
}

# 手动指定目录时，判断它属于哪个客户端
function Resolve-ClientForDir([string]$dir) {
  foreach ($c in $Clients) {
    if (Test-Path -LiteralPath (J2 $dir $c.Exe)) { return $c }
  }
  # 退化判据：卸载程序名是从 exe 名派生的（Uninstall <exe 去扩展名>.exe），不能用 Key 拼
  foreach ($c in $Clients) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($c.Exe)
    if ($base -and (Test-Path -LiteralPath (J2 $dir ('Uninstall ' + $base + '.exe')))) { return $c }
  }
  return $null
}

# 找某个客户端装在哪儿（返回全部命中，去重；已被别的客户端认领的跳过）
function Find-ClientDirs($client, $claimed) {
  $out = New-Object System.Collections.ArrayList
  $seen = @{}
  $add = {
    param($d)
    if (-not $d) { return }
    try { $d = [System.IO.Path]::GetFullPath($d) } catch { return }
    $k = $d.ToLower()
    if ($seen.ContainsKey($k)) { return }
    if ($claimed.ContainsKey($k)) { return }
    if (Test-ClientDir $d $client) { $seen[$k] = $true; [void]$out.Add($d) }
  }

  # 1) 正在运行的进程（最准）
  $procName = $client.Exe -replace '\.exe$', ''
  try {
    foreach ($p in (Get-Process -Name $procName -ErrorAction SilentlyContinue)) {
      try { & $add ([System.IO.Path]::GetDirectoryName($p.Path)) } catch { }
    }
  } catch { }

  # 2) 常见安装位置
  $bases = @()
  if ($env:LOCALAPPDATA) { $bases += $env:LOCALAPPDATA; $bases += (J2 $env:LOCALAPPDATA 'Programs') }
  if ($env:APPDATA) { $bases += $env:APPDATA }
  if ($env:ProgramFiles) { $bases += $env:ProgramFiles }
  if (${env:ProgramFiles(x86)}) { $bases += ${env:ProgramFiles(x86)} }
  foreach ($b in $bases) {
    foreach ($n in $client.Names) { & $add (J2 $b $n) }
  }
  return $out.ToArray()
}

# 客户端看到的"安装目录里的用户内容"
function Get-UserDirs([string]$install) {
  $bad = New-Object System.Collections.ArrayList
  try {
    $items = Get-ChildItem -LiteralPath $install -Directory -Force -ErrorAction SilentlyContinue
  } catch { return $bad.ToArray() }
  foreach ($it in $items) {
    $n = $it.Name.ToLower()
    if ($KnownDirs -contains $n) { continue }
    if ($n.StartsWith('$')) { continue }
    if ($n.StartsWith('nsis')) { continue }
    $isLink = Test-IsLink $it.FullName
    [void]$bad.Add([pscustomobject]@{ Name = $it.Name; Path = $it.FullName; Link = $isLink; Time = $it.LastWriteTime })
  }
  return $bad.ToArray()
}

function Get-AppVersion([string]$install) {
  try {
    $m = J2 (J2 $install 'resources') 'install-manifest.json'
    if (Test-Path -LiteralPath $m) {
      $o = ([System.IO.File]::ReadAllText($m, [System.Text.Encoding]::UTF8) | ConvertFrom-Json)
      if ($o.appVersion) { return [string]$o.appVersion }
    }
  } catch { }
  try {
    $v = J2 $install 'version'
    if (Test-Path -LiteralPath $v) { return ([System.IO.File]::ReadAllText($v)).Trim() }
  } catch { }
  return '未知'
}

# 附带检查：客户端保存的工作空间路径是否落在安装目录里
function Get-WorkspaceHits([string]$dataDir, [string[]]$installDirs) {
  $hits = New-Object System.Collections.ArrayList
  if (-not $dataDir -or -not (Test-Path -LiteralPath $dataDir)) { return $hits.ToArray() }
  foreach ($f in @('settings.json', 'workspace-state.json', 'workspace-display-names.json')) {
    $p = J2 $dataDir $f
    if (-not (Test-Path -LiteralPath $p)) { continue }
    $txt = $null
    try { $txt = [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8) } catch { continue }
    if (-not $txt) { continue }
    foreach ($inst in $installDirs) {
      if ($txt.ToLower().Contains($inst.ToLower())) {
        [void]$hits.Add([pscustomobject]@{ File = $p; Install = $inst })
        break
      }
    }
  }
  return $hits.ToArray()
}

# 破甲是否处于生效状态（只看环境变量，不判好坏）
function Get-PatchState() {
  $names = @('ACC_PRODUCT_CONFIG_PATH', 'ACC_PRODUCT_CONFIG_V3', 'ACC_PRODUCT_CONFIG_V2', 'ACC_PRODUCT_CONFIG')
  $set = @()
  foreach ($n in $names) {
    $v = ''
    try { $v = [Environment]::GetEnvironmentVariable($n, 'User') } catch { }
    if ($v) { $set += $n }
  }
  return $set
}

# ---------------------------------------------------------------- 开始
$started = Get-Date
SayHead 'WorkBuddy 更新拦截 · 一键诊断'
Say ('  时间：' + $started.ToString('yyyy-MM-dd HH:mm:ss'))
try {
  $os = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
  if ($os) { Say ('  系统：' + $os) }
} catch { }
Say '  本脚本只检查，不改动任何文件。'

$foundClients = New-Object System.Collections.ArrayList
$claimed = @{}

if ($InstallDir) {
  # 手动指定（客户端装在很特别的位置时用）
  try { $InstallDir = [System.IO.Path]::GetFullPath($InstallDir) } catch { }
  if (Test-Path -LiteralPath (J2 (J2 $InstallDir 'resources') 'app.asar')) {
    $c = Resolve-ClientForDir $InstallDir
    if (-not $c) {
      # 认不出是哪个版本也要能诊断：目录里挑一个 exe 当主程序名
      $exeName = ''
      try { $exeName = (Get-ChildItem -LiteralPath $InstallDir -Filter '*.exe' -File -ErrorAction SilentlyContinue | Select-Object -First 1).Name } catch { }
      if (-not $exeName) { $exeName = 'WorkBuddy.exe' }
      $c = [pscustomobject]@{ Key = 'unknown'; Name = '（未能识别版本）'; Exe = $exeName; Data = '.workbuddy'; Names = @(); ProcPattern = '' }
    }
    $claimed[$InstallDir.ToLower()] = $true
    [void]$foundClients.Add([pscustomobject]@{ Client = $c; Dirs = @($InstallDir) })
  } else {
    SayBlank
    Say ('  [X] 指定的目录不像客户端安装目录：' + $InstallDir)
    Say '      （里面没有 resources\app.asar，可能选错了一层）'
    $script:ParamError = $true
  }
} else {
  foreach ($c in $Clients) {
    $dirs = Find-ClientDirs $c $claimed
    if ($dirs -and $dirs.Count -gt 0) {
      foreach ($d in $dirs) { $claimed[$d.ToLower()] = $true }
      [void]$foundClients.Add([pscustomobject]@{ Client = $c; Dirs = $dirs })
    }
  }
}

if ($foundClients.Count -eq 0) {
  SayHead '结果'
  Say '  没有找到已安装的 WorkBuddy / WorkBuddy AI。'
  Say ''
  Say '  请先打开一次客户端（首次启动才会生成文件），再重新运行本脚本。'
  Say '  如果客户端装在很特别的位置，可以把安装目录拖到本脚本上运行。'
} else {
  $anyBlocked = $false
  $anyHint = $false

  foreach ($fc in $foundClients) {
    foreach ($install in $fc.Dirs) {
      SayHead ($fc.Client.Name + '：' + $install)
      Say ('  程序版本：' + (Get-AppVersion $install))
      Say ('  主程序  ：' + (J2 $install $fc.Client.Exe))

      $bad = Get-UserDirs $install
      $blockers = @($bad | Where-Object { -not $_.Link })
      $links = @($bad | Where-Object { $_.Link })

      if ($blockers.Count -eq 0) {
        Say ''
        Say '  [OK] 安装目录里没有多余文件夹，更新不会被这条守卫拦住。'
      } else {
        $anyBlocked = $true
        Say ''
        Say ('  [!!] 发现 ' + $blockers.Count + ' 个文件夹会拦住更新：')
        foreach ($b in $blockers) {
          Say ('        · ' + $b.Name)
          Say ('          路径：' + $b.Path)
          Say ('          时间：' + $b.Time.ToString('yyyy-MM-dd HH:mm'))
        }
      }

      if ($links.Count -gt 0) {
        Say ''
        Say ('  [提示] 另有 ' + $links.Count + ' 个是链接目录（junction/符号链接），不会触发拦截：')
        foreach ($l in $links) { Say ('        · ' + $l.Name) }
      }

      # 附带检查：工作空间配置
      $dataDir = J2 $env:USERPROFILE $fc.Client.Data
      $wh = Get-WorkspaceHits $dataDir @($install)
      if ($wh.Count -gt 0) {
        $anyHint = $true
        Say ''
        Say '  [注意] 客户端保存的工作空间路径指向了安装目录：'
        foreach ($h in $wh) { Say ('        · ' + $h.File) }
        Say '        这会让客户端往安装目录里创建项目文件夹，然后被更新守卫拦住。'
      }
    }
  }

  # 破甲状态
  SayHead '其它信息'
  $pset = Get-PatchState
  if ($pset.Count -gt 0) {
    Say ('  破甲环境变量：已设置（' + ($pset -join ', ') + '）')
  } else {
    Say '  破甲环境变量：未设置'
  }

  # 结论
  SayHead '结论'
  if ($anyBlocked) {
    Say '  [!!] 更新会被拦住 —— 上面列出的文件夹就是原因。'
    Say ''
    Say '  怎么处理（三步）：'
    Say '    1) 把上面列出的文件夹【整个剪切】到别的地方，'
    Say '       例如  D:\WorkBuddy项目\  ，不要只删里面的文件。'
    Say '    2) 打开客户端，如果里面有工作空间指向这些文件夹，改成新位置。'
    Say '    3) 完全退出客户端再打开，然后重新点「检查更新」。'
    Say ''
    Say '  注意：请用"剪切/移动"，不要直接删除，避免丢数据。'
  } elseif ($anyHint) {
    Say '  [注意] 安装目录本身是干净的，但工作空间配置指向了安装目录，'
    Say '         下次客户端启动可能又会在里面建文件夹。建议先把工作空间改到别处。'
  } else {
    Say '  [OK] 没有发现会导致更新被拦的问题。'
    Say ''
    Say '  如果更新仍然失败，把本报告发给提供者，让他看客户端日志：'
    foreach ($fc in $foundClients) {
      Say ('    ' + (J2 (J2 $env:USERPROFILE $fc.Client.Data) 'logs'))
    }
  }
}

# ---------------------------------------------------------------- 写报告
$reportText = ($script:Report -join [Environment]::NewLine)

if (-not $ReportPath) {
  $cands = @()
  $cands += (J2 $env:USERPROFILE 'Desktop\更新拦截诊断报告.txt')
  if ($PSScriptRoot) { $cands += (J2 $PSScriptRoot '更新拦截诊断报告.txt') }
  $cands += (J2 $env:TEMP '更新拦截诊断报告.txt')
  foreach ($c in $cands) {
    try {
      $dir = [System.IO.Path]::GetDirectoryName($c)
      if ($dir -and -not (Test-Path -LiteralPath $dir)) { continue }
      [System.IO.File]::WriteAllText($c, $reportText, (New-Object System.Text.UTF8Encoding($true)))
      $ReportPath = $c
      break
    } catch { }
  }
} else {
  try {
    [System.IO.File]::WriteAllText($ReportPath, $reportText, (New-Object System.Text.UTF8Encoding($true)))
  } catch { $ReportPath = '' }
}

if ($ReportPath) {
  Write-Host ''
  Write-Host ('  报告已保存：' + $ReportPath) -ForegroundColor DarkGray
}

# ---------------------------------------------------------------- 弹窗
$verdictTitle = 'WorkBuddy 更新拦截诊断'
$verdictBody = ''
if ($foundClients.Count -eq 0) {
  $verdictBody = "没有找到已安装的 WorkBuddy / WorkBuddy AI。`n`n请先打开一次客户端，再重新运行本脚本。"
} elseif ($anyBlocked) {
  $lines = @()
  foreach ($fc in $foundClients) {
    foreach ($install in $fc.Dirs) {
      foreach ($b in (Get-UserDirs $install)) {
        if (-not $b.Link) { $lines += ('· ' + $b.Path) }
      }
    }
  }
  $verdictBody = "发现会导致更新被拦的文件夹：`n`n" + (($lines | Select-Object -First 8) -join "`n")
  if ($lines.Count -gt 8) { $verdictBody += "`n… 还有 " + ($lines.Count - 8) + " 个，见报告文件。" }
  $verdictBody += "`n`n处理办法：把这些文件夹【剪切】到别的地方（例如 D:\WorkBuddy项目\），`n然后完全退出客户端再打开，重新检查更新。`n`n注意：移动，不要删除。`n`n点「是」打开安装目录，方便你剪切。"
} elseif ($anyHint) {
  $verdictBody = "安装目录本身是干净的。`n`n但客户端保存的工作空间路径指向了安装目录，`n建议先把工作空间改到别的地方（例如 D:\WorkBuddy项目\）。"
} else {
  $verdictBody = "没有发现会导致更新被拦的问题。`n`n如果更新还是失败，请把报告文件发给提供者。"
}
if ($ReportPath) { $verdictBody += "`n`n报告文件：`n" + $ReportPath }

$openDir = ''
if ($anyBlocked) {
  foreach ($fc in $foundClients) { if ($fc.Dirs.Count -gt 0) { $openDir = $fc.Dirs[0]; break } }
}

$answer = 2
try {
  Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
  $icon = [System.Windows.Forms.MessageBoxIcon]::Warning
  if (-not $anyBlocked -and -not $anyHint) { $icon = [System.Windows.Forms.MessageBoxIcon]::Information }
  $btns = [System.Windows.Forms.MessageBoxButtons]::OK
  if ($openDir) { $btns = [System.Windows.Forms.MessageBoxButtons]::YesNo }
  $r = [System.Windows.Forms.MessageBox]::Show($verdictBody, $verdictTitle, $btns, $icon)
  if ($r -eq [System.Windows.Forms.DialogResult]::Yes) { $answer = 1 } else { $answer = 0 }
} catch {
  try {
    $ws = New-Object -ComObject WScript.Shell
    $r = $ws.Popup($verdictBody, 0, $verdictTitle, 64)
    if ($r -eq 1 -and $openDir) { $answer = 1 } else { $answer = 0 }
  } catch { $answer = 0 }
}

if ($answer -eq 1 -and $openDir) {
  try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $openDir + '"') } catch { }
}

if ($Host.Name -eq 'ConsoleHost' -and -not $NoPause) {
  Write-Host ''
  Write-Host '  按回车键关闭本窗口...' -ForegroundColor DarkGray
  [void](Read-Host)
}

# ---------------------------------------------------------------- 退出码契约（v2.0）
#   0 = 未发现会挡住更新的文件夹
#   1 = 发现了拦截项（需要用户按报告处理）
#   2 = 没找到任何客户端
#   3 = 参数/环境问题（例如 -InstallDir 指错了）
if ($script:ParamError) { exit 3 }
if ($foundClients.Count -eq 0) { exit 2 }
if ($anyBlocked) { exit 1 }
exit 0
