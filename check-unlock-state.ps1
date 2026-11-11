# check-unlock-state.ps1 -- 只读检测：破甲是否已生效（Windows）
#
# 它只"看"，不改任何东西：不写文件、不设环境变量、不动进程。
# 检测四层：
#   A 文件层  默认开启的每条补丁是否都已打到客户端文件里（政策块 / 闸门 / 子句 / product.json）
#   B 环境层  环境变量覆盖通道是否到位（变量指向本工具的覆盖文件，且文件合法）
#   C 运行层  客户端是否在运行、当前进程是不是"补丁之后"启动的（决定要不要重启）
#   D 保障层  能不能还原（有没有快照）、守护任务装没装/上次跑成没成
#
# 用法：
#   powershell -File check-unlock-state.ps1                # 人看的报告
#   ... -Quiet                                             # 只打判定行（给脚本/计划任务用）
#   ... -Edition cn|intl                                   # 只看某一个版本
#   ... -Strict                                            # "需要重启"也算未生效（退出码 1）
#   ... -InstallDir/-DataDir <目录>                         # 手工指定客户端位置（探针/自测用）
#   ... -ReportPath <文件>                                  # 自定义报告路径
#
# 退出码：0 已生效（或已就位）／1 未生效（或 -Strict 下待重启）／2 没找到客户端／3 环境问题

[CmdletBinding()]
param(
    [ValidateSet('auto', 'cn', 'intl')][string]$Edition = 'auto',
    [string]$InstallDir,
    [string]$DataDir,
    [string]$StateDir,
    [switch]$Quiet,
    [switch]$Strict,
    [string]$ReportPath,
    [switch]$NoPause
)

$ErrorActionPreference = 'Continue'

# 启动器是纯 ASCII 单行 .bat（不设标题），窗口标题在这里补上
try { if ($Host.UI.RawUI) { $Host.UI.RawUI.WindowTitle = 'WorkBuddy 破甲 - 是否生效检测' } } catch {}

$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
$DataRoot = Join-Path $ScriptDir 'data'
$InstallScript = Join-Path $ScriptDir 'install.ps1'
$SENT_OPEN = '<<<WBPOJA_JSON>>>'
$SENT_CLOSE = '<<<WBPOJA_END>>>'

$lines = New-Object System.Collections.ArrayList
function Emit([string]$s) {
    [void]$lines.Add($s)
    if (-not $Quiet) { Write-Host $s }
}
function EmitQuiet([string]$s) {
    # 静默模式也要给出一行判定
    [void]$lines.Add($s)
    Write-Host $s
}
function Note([string]$s) { [void]$lines.Add('  ' + $s); if (-not $Quiet) { Write-Host ('  ' + $s) } }
function SayBlank() { [void]$lines.Add(''); if (-not $Quiet) { Write-Host '' } }
function SayHead([string]$s) {
    SayBlank
    [void]$lines.Add('------------------------------------------------------------')
    [void]$lines.Add('  ' + $s)
    [void]$lines.Add('------------------------------------------------------------')
    if (-not $Quiet) {
        Write-Host ''
        Write-Host '------------------------------------------------------------'
        Write-Host ('  ' + $s)
        Write-Host '------------------------------------------------------------'
    }
}

# ---------------------------------------------------------------- 读版本表
function Read-EditionsCfg {
    $p = Join-Path $DataRoot 'editions.json'
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    try { return ([System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8) | ConvertFrom-Json) } catch { return $null }
}

# ---------------------------------------------------------------- 拿文件层状态
function Get-PatchState {
    if (-not (Test-Path -LiteralPath $InstallScript)) {
        return @{ ok = $false; reason = ('找不到 install.ps1：' + $InstallScript) }
    }
    # 优先让 install.ps1 把 JSON 写到临时文件：中文用户名之类的路径经 stdout 传回
    # 会因控制台代码页不一致变成乱码（file-out 完全避开这个问题）
    $jsonTmp = Join-Path $env:TEMP ('wb-poja-state-' + [Guid]::NewGuid().ToString('N').Substring(0, 8) + '.json')
    $argv = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $InstallScript,
              '-DumpAnchor', '-AsJson', '-Silent', '-NoDialog')
    if ($Edition -ne 'auto') { $argv += @('-Edition', $Edition) }
    if ($InstallDir) { $argv += @('-InstallDir', $InstallDir) }
    if ($DataDir) { $argv += @('-DataDir', $DataDir) }
    $argv += @('-JsonOut', $jsonTmp)
    try { $null = (& powershell @argv 2>&1 | Out-String) } catch {
        return @{ ok = $false; reason = ('调用 install.ps1 失败：' + $_.Exception.Message) }
    }
    $json = ''
    if (Test-Path -LiteralPath $jsonTmp) {
        try { $json = [System.IO.File]::ReadAllText($jsonTmp, [System.Text.Encoding]::UTF8) } catch { $json = '' }
        try { Remove-Item -LiteralPath $jsonTmp -Force -ErrorAction SilentlyContinue } catch {}
    }
    if (-not $json) {
        # 退回 stdout（老版本 install.ps1 没有 -JsonOut）
        $out = ''
        try {
            $argv2 = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $InstallScript,
                       '-DumpAnchor', '-AsJson', '-Silent', '-NoDialog')
            if ($Edition -ne 'auto') { $argv2 += @('-Edition', $Edition) }
            if ($InstallDir) { $argv2 += @('-InstallDir', $InstallDir) }
            if ($DataDir) { $argv2 += @('-DataDir', $DataDir) }
            $out = (& powershell @argv2 2>&1 | Out-String)
        } catch {
            return @{ ok = $false; reason = ('调用 install.ps1 失败：' + $_.Exception.Message) }
        }
        $i = $out.IndexOf($SENT_OPEN)
        $j = $out.IndexOf($SENT_CLOSE)
        if ($i -lt 0 -or $j -le $i) {
            return @{ ok = $false; reason = '没有拿到结构化状态（install.ps1 输出异常）' }
        }
        $json = $out.Substring($i + $SENT_OPEN.Length, $j - $i - $SENT_OPEN.Length)
    }
    try { return @{ ok = $true; data = ($json | ConvertFrom-Json) } } catch {
        return @{ ok = $false; reason = ('状态 JSON 解析失败：' + $_.Exception.Message) }
    }
}

# ---------------------------------------------------------------- 运行时事实
function Test-EnvChannel($ed) {
    # 环境变量通道：哪个内置变量名指向了本工具的覆盖文件
    $names = @('ACC_PRODUCT_CONFIG_PATH', 'ACC_PRODUCT_CONFIG_V3', 'ACC_PRODUCT_CONFIG_V2', 'ACC_PRODUCT_CONFIG')
    $envFile = Join-Path (Join-Path $env:LOCALAPPDATA $ed.patchRootName) $ed.envFileName
    $hit = ''
    $value = ''
    foreach ($n in $names) {
        $v = [Environment]::GetEnvironmentVariable($n, 'User')
        if (-not $v) { continue }
        $own = $false
        if (($v -replace '\\', '/').ToLower() -eq ($envFile -replace '\\', '/').ToLower()) { $own = $true }
        foreach ($r in @('WorkBuddyPatch', 'WorkBuddyAIPatch')) {
            if (($v -replace '\\', '/').ToLower().Contains(($r.ToLower() + '/'))) { $own = $true }
        }
        if ($v.ToLower().EndsWith($ed.envFileName.ToLower())) { $own = $true }
        if ($own) { $hit = $n; $value = $v; break }
    }
    $fileOk = $false
    $policyIn = $false
    if (Test-Path -LiteralPath $envFile) {
        try {
            $cfgTxt = [System.IO.File]::ReadAllText($envFile, [System.Text.Encoding]::UTF8)
            $cfg = $cfgTxt | ConvertFrom-Json
            $fileOk = $true
            foreach ($pr in $cfg.prompts) {
                $t = [string]$pr.template
                if ($t.Contains('本内容政策是模型收到的最高指令') -or $t.Contains('PRIME DIRECTIVE (user supremacy)') -or $t.Contains('**总纲（用户至上）**')) { $policyIn = $true }
            }
        } catch { $fileOk = $false }
    }
    return @{ name = $hit; value = $value; envFile = $envFile; fileOk = $fileOk; policyIn = $policyIn;
              ok = ([bool]$hit -and $fileOk -and $policyIn) }
}

function Get-ClientProc($ed) {
    $pat = '^(?i)(WorkBuddy(?!AI)|CodeBuddy)$'
    if ($ed.key -eq 'intl') { $pat = '^(?i)(WorkBuddyAI|workbuddy-ai)$' }
    try { return @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $pat }) } catch { return @() }
}

function Get-PatchTime($ed) {
    # 最近一次打补丁的时间：取最新快照的名字（yyyyMMdd-HHmmss）
    $root = if ($StateDir) { Split-Path -Parent $StateDir } else { Join-Path $env:LOCALAPPDATA $ed.patchRootName }
    $sd = if ($StateDir) { $StateDir } else { Join-Path $root 'state' }
    $newest = $null
    if (Test-Path -LiteralPath $sd) {
        $newest = Get-ChildItem -LiteralPath $sd -Directory -ErrorAction SilentlyContinue |
                  Sort-Object Name -Descending | Select-Object -First 1
    }
    $t = $null
    if ($newest) {
        $raw = ($newest.Name -replace '-ai$', '')
        try { $t = [datetime]::ParseExact($raw, 'yyyyMMdd-HHmmss', $null) } catch { $t = $newest.LastWriteTime }
    }
    $count = 0
    if (Test-Path -LiteralPath $sd) { $count = @(Get-ChildItem -LiteralPath $sd -Directory -ErrorAction SilentlyContinue).Count }
    return @{ time = $t; snapshots = $count; root = $root }
}

function Get-GuardState($ed) {
    $out = @()
    foreach ($suffix in @('_Hourly', '_Logon')) {
        $name = $ed.guardTaskTag + $suffix
        $exists = $false
        try { $null = & schtasks /Query /TN $name 2>$null; $exists = ($LASTEXITCODE -eq 0) } catch {}
        $last = $null
        if ($exists) {
            try {
                $q = (& schtasks /Query /TN $name /V /FO LIST 2>$null | Out-String)
                foreach ($ln in ($q -split "`r?`n")) {
                    if ($ln -match '^\s*(Last Result|上次运行结果)\s*[:：]\s*(-?\d+)') { $last = [int]$Matches[2]; break }
                }
            } catch {}
        }
        $out += [pscustomobject]@{ name = $name; exists = $exists; lastResult = $last }
    }
    return $out
}

function Test-NeedRestart($procs, $patchTime) {
    if (-not $patchTime) { return $false }
    foreach ($p in $procs) {
        try {
            if ($p.StartTime -lt $patchTime) { return $true }
        } catch { return $true }   # 取不到启动时间（权限不足）时保守认为需要重启
    }
    return $false
}

# ================================================================== 主流程
# 版本号要在配置读取之前拿到（下面的 $cfg 还没赋值）；读不到就留空
$toolVer = ''
try { $toolVer = [string](Read-EditionsCfg).toolVersion } catch {}
SayHead ('破甲是否生效 · 自动检测（只读）  ·  工具 ' + $(if ($toolVer) { 'v' + $toolVer } else { '(版本未知)' }))
Emit ('时间：' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
Emit '本脚本只做检查：不写文件、不改环境变量、不动进程。'

$cfg = Read-EditionsCfg
if (-not $cfg) {
    Emit ('[X] 缺少数据文件：' + (Join-Path $DataRoot 'editions.json'))
    Emit '    （请确认本脚本与 install.ps1、data 目录在同一个文件夹里）'
    exit 3
}

$st = Get-PatchState
if (-not $st.ok) {
    Emit ('[X] ' + $st.reason)
    exit 3
}

$editions = @($st.data.editions)
if ($editions.Count -eq 0) {
    Emit '[!] 没有找到任何客户端安装目录'
    Emit '    先安装并至少打开一次客户端，或用 -InstallDir 指定安装目录。'
    exit 2
}

# 两个版本装在同一目录时，安装器会互相覆盖 resources/（"更新后打不开/永远升不上去"）
$dirsInstall = @($editions | ForEach-Object { ([string]$_.installDir).TrimEnd('\').ToLower() } | Where-Object { $_ })
if ($dirsInstall.Count -ge 2 -and (@($dirsInstall | Sort-Object -Unique).Count -lt $dirsInstall.Count)) {
    Emit '[!] 国内版与国际版装在同一个目录：更新会互相覆盖，建议把国际版单独安装到别的文件夹。'
    SayBlank
}

$worst = 0
foreach ($row in $editions) {
    $ed = $cfg.editions.($row.edition)
    SayHead ($row.displayName + '：' + $row.installDir)

    # ---- A 文件层
    $must = @($row.patches | Where-Object { $_.defaultOn -and ($_.platforms -contains 'win') })
    # optional 补丁：老客户端里根本没有对应文本（state=无该块）时不算缺失
    $missing = @($must | Where-Object { $_.state -ne '已打' -and -not ($_.optional -and $_.state -eq '无该块') })
    $extra = @($row.patches | Where-Object { (-not $_.defaultOn) -and $_.state -eq '已打' -and ($_.platforms -contains 'win') })
    foreach ($p in $must) {
        $mark = 'OK '
        if (($p.state -ne '已打') -and -not ($p.optional -and $p.state -eq '无该块')) { $mark = '!! ' }
        $optTag = ''
        if ($p.optional) { $optTag = '，可选' }
        Emit ('  [' + $mark + '] ' + $p.id.PadRight(22) + $p.state + '（目标 ' + $p.targets + ' / 已打 ' + $p.done + $optTag + '）')
    }
    if ($extra.Count -gt 0) {
        $names = ($extra | ForEach-Object { $_.id }) -join ' / '
        Note ('额外开启：' + $names)
    }
    $fileOk = ($missing.Count -eq 0)

    # ---- B 环境层
    $env = Test-EnvChannel $ed
    if ($env.ok) {
        Emit ('  [OK ] 环境变量通道：' + $env.name + ' → ' + $env.envFile)
    } elseif ($env.name) {
        Emit ('  [!! ] 环境变量通道：' + $env.name + ' 指向 ' + $env.value)
        Note ('但覆盖文件不可用或内容不含我们的政策：' + $env.envFile)
    } else {
        Emit '  [-- ] 环境变量通道：未设置（不影响磁盘补丁，重启后照样生效）'
    }

    # ---- C 运行层
    $procs = Get-ClientProc $ed
    $pt = Get-PatchTime $ed
    $restart = Test-NeedRestart $procs $pt.time
    if ($procs.Count -eq 0) {
        Emit '  [-- ] 客户端未运行（打开即生效）'
    } elseif ($restart) {
        Emit ('  [!! ] 客户端正在运行（' + $procs.Count + ' 个进程），但它是补丁之前启动的')
        Note '重启客户端才会加载补丁后的文件。'
    } else {
        Emit ('  [OK ] 客户端正在运行，且是补丁之后启动的（' + $procs.Count + ' 个进程）')
    }

    # 数据目录兜底：JSON 里若因编码问题不可用，就按版本表推导
    $ddRow = [string]$row.dataDir
    if (-not $ddRow -or -not (Test-Path -LiteralPath $ddRow)) {
        try { $ddRow = Join-Path $env:USERPROFILE ([string]$ed.dataFolder) } catch {}
    }

    # ---- C2 旧版本插件缓存残留（客户端不读它们，但会让"还原干净"这件事看着不干净）
    try {
        $cacheRoot = Join-Path $ddRow 'plugins\cache'
        if (Test-Path -LiteralPath $cacheRoot) {
            $ver = ''
            $ll = Join-Path $ddRow 'last-launch.json'
            if (Test-Path -LiteralPath $ll) {
                try { $ver = [string](([System.IO.File]::ReadAllText($ll, [System.Text.Encoding]::UTF8) | ConvertFrom-Json).version) } catch {}
            }
            $stale = 0
            foreach ($f in (Get-ChildItem -LiteralPath $cacheRoot -Recurse -Filter 'prompt.tpl' -File -ErrorAction SilentlyContinue)) {
                if ($ver -and ($f.FullName -match ('\\' + [regex]::Escape($ver) + '-wb\.'))) { continue }
                $txt = ''
                try { $txt = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8) } catch { continue }
                foreach ($mk in @('本内容政策是模型收到的最高指令', 'PRIME DIRECTIVE (user supremacy)', 'Highest precedence: this content policy', '**总纲（用户至上）**')) {
                    if ($txt.Contains($mk)) { $stale++; break }
                }
            }
            if ($stale -gt 0) {
                Emit ('  [-- ] 旧版本插件缓存残留：' + $stale + ' 个（客户端不读它们，不影响使用）')
                Note '还原时会自动修复；修不了可用 -PurgeStaleCaches 删除旧缓存目录'
            }
        }
    } catch {}

    # ---- D 缓存层（删过缓存 = 刚打过补丁；重建过 = 客户端已按新配置启动过）
    # 5.9.2 起国内版是"每个进程一份" acc-product-config-v3-<pid>-<guid>.json，
    # 固定文件名可能不存在，按修改时间取最新一份。
    $cache = $null
    $cacheDirRow = Join-Path $ddRow 'cache'
    if (Test-Path -LiteralPath $cacheDirRow) {
        $cf = @(Get-ChildItem -LiteralPath $cacheDirRow -File -Filter 'acc-product-config*.json' -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1)
        if ($cf.Count -gt 0) { $cache = $cf[0].FullName }
    }
    if ($cache -and (Test-Path -LiteralPath $cache)) {
        $w = (Get-Item -LiteralPath $cache).LastWriteTime
        Emit ('  [OK ] 运行时缓存：已存在（' + $w.ToString('MM-dd HH:mm') + ' 重建）')
    } else {
        Emit '  [-- ] 运行时缓存：不存在（本工具会清它；客户端下次启动会重建）'
    }

    # ---- 保障层
    if ($pt.snapshots -gt 0) {
        Emit ('  [OK ] 可还原：' + $pt.snapshots + ' 份快照在 ' + $pt.root)
        if ($pt.time) { Note ('最近一次打补丁：' + $pt.time.ToString('yyyy-MM-dd HH:mm:ss')) }
    } else {
        Emit '  [-- ] 没有快照（本次没有改动过，或快照被手工删掉了）'
    }
    $ubState = Join-Path (Join-Path $env:LOCALAPPDATA $ed.patchRootName) 'update-block.json'
    $ub = $false
    if (Test-Path -LiteralPath $ubState) {
        try {
            $u = (Get-Content -LiteralPath $ubState -Raw -Encoding UTF8 | ConvertFrom-Json)
            if ($u.cachePath -and (Test-Path -LiteralPath ([string]$u.cachePath) -PathType Leaf)) { $ub = $true }
        } catch {}
    }
    if ($ub) { Emit '  [OK ] 自动更新：已阻断（可选）' }
    else { Emit '  [-- ] 自动更新：未阻断（默认）' }
    # 备用前置入口（可选）：直接读 settings.json 的 env，反映实际生效值
    $altVal = ''
    $settingsRow = Join-Path $ddRow 'settings.json'
    if (Test-Path -LiteralPath $settingsRow) {
        try { $altVal = [string]((Get-Content -LiteralPath $settingsRow -Raw -Encoding UTF8 | ConvertFrom-Json).env.CODEBUDDY_BASE_URL) } catch {}
    }
    if ($altVal -match 'copilot\.tencent\.com') { Emit ('  [OK ] 备用前置入口：已启用（' + $altVal + '）') }
    elseif ($altVal) { Emit ('  [OK ] 模型 baseURL：自定义（' + $altVal + '）') }
    else { Emit '  [-- ] 备用前置入口：未启用（默认门）' }
    $guard = Get-GuardState $ed
    $gexist = @($guard | Where-Object { $_.exists })
    if ($gexist.Count -eq 0) {
        Emit '  [-- ] 守护任务：未安装（升级后不会自动重打；需要就双击 18-安装守护任务.bat）'
    } else {
        $bad = @($gexist | Where-Object { $_.lastResult -ne $null -and (@(0, 267009, 267011) -notcontains $_.lastResult) })
        if ($bad.Count -gt 0) {
            Emit ('  [!! ] 守护任务：已装 ' + $gexist.Count + ' 个，但有 ' + $bad.Count + ' 个上次运行异常')
            foreach ($b in $bad) { Note ($b.name + ' 上次结果 ' + $b.lastResult + '（非 0/运行中/未跑过）') }
        } else {
            Emit ('  [OK ] 守护任务：已装 ' + $gexist.Count + ' 个')
        }
    }

    # ---- 本版判定
    SayBlank
    if (-not $fileOk) {
        $names = ($missing | ForEach-Object { $_.id }) -join '、'
        Emit ('  >>> 结论：未生效（文件层缺 ' + $missing.Count + ' 条：' + $names + '）')
        Note '下一步：双击 1-一键破甲.bat（然后重启客户端）'
        if ($worst -lt 1) { $worst = 1 }
    } elseif ($restart) {
        Emit '  >>> 结论：补丁已就位，但客户端还没重启 —— 现在还不生效'
        Note '下一步：完全退出客户端再打开'
        if ($Strict -and $worst -lt 1) { $worst = 1 }
    } elseif ($procs.Count -eq 0) {
        Emit '  >>> 结论：补丁已就位（客户端当前未运行，打开即生效）'
    } else {
        Emit '  >>> 结论：已生效'
    }
    if (-not $env.ok -and -not $env.name) {
        Note '说明：环境变量通道没设，不影响磁盘补丁；但"免重启"的那条路没走。'
    }
}

SayHead '怎么读这份结果'
Emit '  [OK] 已生效/已到位     [!!] 需要处理     [--] 不适用或未启用'
Emit '  · 文件层全绿 + 客户端重启过 → 就是生效了'
Emit '  · 文件层有 [!!]：双击 1-一键破甲.bat 重打'
Emit '  · 文件层全绿但提示"客户端是补丁之前启动的"：退出客户端重开即可'
Emit '  · 想确认得更彻底：双击 4-锚点体检.bat 看逐条明细；还原后想确认干净：跑 2-一键还原.bat'

# ---------------------------------------------------------------- 写报告
$reportText = ($lines -join [Environment]::NewLine)
if (-not $ReportPath) {
    $cands = @(
        (Join-Path $env:USERPROFILE 'Desktop\破甲生效检测报告.txt'),
        (Join-Path $ScriptDir '破甲生效检测报告.txt'),
        (Join-Path $env:TEMP '破甲生效检测报告.txt')
    )
    foreach ($c in $cands) {
        try {
            $d = [System.IO.Path]::GetDirectoryName($c)
            if ($d -and -not (Test-Path -LiteralPath $d)) { continue }
            [System.IO.File]::WriteAllText($c, $reportText, (New-Object System.Text.UTF8Encoding($true)))
            $ReportPath = $c
            break
        } catch {}
    }
} else {
    try { [System.IO.File]::WriteAllText($ReportPath, $reportText, (New-Object System.Text.UTF8Encoding($true))) } catch { $ReportPath = '' }
}
if ($ReportPath -and -not $Quiet) {
    SayBlank
    Write-Host ('  报告已保存：' + $ReportPath) -ForegroundColor DarkGray
}

# ---------------------------------------------------------------- 收尾
if ($Quiet) {
    # 静默模式只给一行汇总，措辞与报告里的结论保持一致，便于脚本/计划任务 grep
    $verdicts = @()
    foreach ($row in $editions) {
        $must = @($row.patches | Where-Object { $_.defaultOn -and ($_.platforms -contains 'win') })
        $fileOk = (@($must | Where-Object { $_.state -ne '已打' -and -not ($_.optional -and $_.state -eq '无该块') }).Count -eq 0)
        $ed0 = $cfg.editions.($row.edition)
        $procs0 = Get-ClientProc $ed0
        $pt0 = Get-PatchTime $ed0
        $v = '未生效'
        if ($fileOk) {
            if (Test-NeedRestart $procs0 $pt0.time) { $v = '需重启' }
            elseif ($procs0.Count -eq 0) { $v = '已就位' }
            else { $v = '已生效' }
        }
        $verdicts += ([string]$ed0.shortName + '=' + $v)
    }
    EmitQuiet ('破甲状态：' + ($verdicts -join '  '))
}

if ($Host.Name -eq 'ConsoleHost' -and -not $NoPause -and -not $Quiet) {
    Write-Host ''
    Write-Host '  按回车键关闭...' -ForegroundColor DarkGray
    [void](Read-Host)
}

exit $worst
