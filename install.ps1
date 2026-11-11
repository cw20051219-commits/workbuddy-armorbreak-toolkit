# install.ps1 -- WorkBuddy 破甲工具包（Windows / 一个包同时支持国内版 + 国际版）
# 版本号只在 data/editions.json 的 toolVersion 里定义，这里不写死
#
# 与 v1.x 的关键差别：
#   1) 版本表与补丁表外置在 data/（editions.json / patches/*.json），代码不硬编码路径与锚点。
#      客户端升级把目标文件搬家后，只改数据文件，不改代码。
#   2) 默认 -Edition auto：检测到几个客户端就处理几个，报告分开列。
#   3) 身份校验双向：国际版要求命中 workbuddy.ai，国内版要求"不是国际版"——两版同装也不互相误伤。
#   4) 环境变量一律 owned-only：只清"本工具族写过"的值，别名指向别人时不动；写前记录原值。
#   5) 每次写盘登记 pending，收尾逐字节回读复查（防"工具说成功其实没生效"）。
#   6) 还原同时扫描国内版与国际版两个备份根，v1.x 的老快照照样能还原。
#
# 兼容 Windows PowerShell 5.1 与 PowerShell 7+；支持路径含空格/中文。

[CmdletBinding()]
param(
    [ValidateSet('Patch', 'Restore')][string]$Mode = 'Patch',
    [ValidateSet('auto', 'cn', 'intl')][string]$Edition = 'auto',
    [ValidateSet('standard', 'strong')][string]$Level = 'standard',
    [ValidateSet('auto', 'en', 'zh')][string]$PolicyLang = 'auto',
    [string]$InstallDir,
    [string]$DataDir,
    [string]$StateDir,
    [ValidateSet('a', 'ab')][string]$Channel = 'a',   # mac 专用；Windows 忽略
    [switch]$NoEnv,
    [switch]$NoDisk,
    [switch]$NoRestart,
    [switch]$NoDialog,
    [switch]$DryRun,
    [switch]$Elevated,
    [switch]$Silent,
    [switch]$WebFilter,        # 02 网页过滤解除（默认关）
    [switch]$NoFileGuard,      # 03 文件保护中和（默认关）
    [switch]$AskMode,          # 08 Ask 模式解除（默认关）
    [switch]$MinimalPrompt,    # 10 整段替换 cli-agent-prompt 为精简提示（默认关）
    [switch]$EngineeringPrompt,      # 11 整段替换为工程助手风格提示（无政策块，默认关）
    [switch]$ContractPrompt,    # 12 整段替换为行为契约式提示（无政策块，默认关）
    [switch]$BlockUpdate,      # 可选：阻断客户端自动更新
    [switch]$UnblockUpdate,    # 恢复自动更新
    [switch]$AltFrontDoor,     # 可选：内置模型改走备用前置入口（copilot.tencent.com）
    [switch]$AltFrontDoorOff,  # 关闭备用入口，恢复默认门
    [string]$UpdateCachePath,  # 自测用：覆盖更新缓存路径
    [switch]$KeepCache,        # 05 不清理运行时缓存
    [switch]$Guard,
    [switch]$GuardRemove,
    [switch]$GuardStatus,
    [switch]$CloudMem,
    [switch]$CloudMemRevert,
    [switch]$CloudMemStatus,
    [switch]$NoJsGate,         # 跳过 JS 命令闸门解锁（客户端版本敏感时可关掉）
    [switch]$DeepClean,        # 没有干净备份时：按内容把可逆补丁直接摘掉
    [switch]$PurgeStaleCaches, # 删除"旧版本插件缓存"里残留了我们改动的缓存目录
    [switch]$DumpAnchor,       # 锚点体检（只读）
    [switch]$AsJson,           # 配合 -DumpAnchor：输出机器可读的 JSON
    [string]$JsonOut,          # 配合 -DumpAnchor -AsJson：把 JSON 写到文件（UTF-8，避免控制台代码页把中文路径搞乱）
    [switch]$Help
)

$ErrorActionPreference = 'Continue'

# ------------------------------------------------------------------ 基本路径
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
$DataRoot  = Join-Path $ScriptDir 'data'
$IsSilent  = [bool]$Silent
$Platform  = 'win'
$script:LogFile = $null
$script:Pending = New-Object System.Collections.ArrayList
$script:EnabledPatchIds = New-Object System.Collections.ArrayList
$script:StaleSkipped = 0
$script:SnapDir = $null

# ------------------------------------------------------------------ 小工具
function Write-Log([string]$msg, [string]$color) {
    if (-not $IsSilent) {
        if ($color) { try { Write-Host $msg -ForegroundColor $color } catch { Write-Host $msg } }
        else { Write-Host $msg }
    }
    if ($script:LogFile) {
        try {
            Add-Content -LiteralPath $script:LogFile -Value ('[' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] ' + $msg) -Encoding UTF8
        } catch {}
    }
}

function Show-Message([string]$title, [string]$text) {
    if ($NoDialog -or $IsSilent) { return }
    try {
        $ws = New-Object -ComObject WScript.Shell
        $ws.Popup($text, 0, $title, 64) | Out-Null
    } catch {
        try {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show($text, $title) | Out-Null
        } catch {}
    }
}

function Read-Json([string]$path) {
    return (([System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)) | ConvertFrom-Json)
}

function ConvertTo-PathSep([string]$rel) { return ($rel -replace '/', '\') }

function Get-FileEol([string]$text) {
    $crlf = ([regex]::Matches($text, "`r`n")).Count
    $lf   = ([regex]::Matches($text, "`n")).Count
    if ($lf -eq 0) { return "`n" }
    if ($crlf -eq 0) { return "`n" }
    if ($crlf -eq $lf) { return "`r`n" }
    return "`n"
}

function ConvertTo-Nl([string]$text, [string]$eol) {
    $t = $text -replace "`r`n", "`n"
    $t = $t -replace "`r", "`n"
    if ($eol -eq "`r`n") { return ($t -replace "`n", "`r`n") }
    return $t
}

function Get-JsonBody([string]$s) {
    $t = $s -replace '\\', '\\'
    $t = $t -replace '"', '\"'
    $t = $t -replace "`r`n", '\n'
    $t = $t -replace "`n", '\n'
    $t = $t -replace "`r", ''
    $t = $t -replace "`t", '\t'
    return $t
}

function Test-Writable([string]$dir) {
    if (-not (Test-Path -LiteralPath $dir)) { return $false }
    try {
        $probe = Join-Path $dir ('.wb-write-test-' + [Guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($probe, 'x')
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return $true
    } catch { return $false }
}

# 只读/ACL 场景：Test-Writable 只证明“能新建文件”，不能证明“能覆盖已有只读文件”。
# 这里再打开第一个 .tpl 做写测试；只读位单独判，避免把“被客户端占用”误判成权限问题。
function Test-FileWritable([string]$file) {
    try {
        $fs = [System.IO.File]::Open($file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $fs.Close()
        return $true
    } catch [System.UnauthorizedAccessException] { return $false } catch { return $true }
}
function Test-DirPatchable([string]$dir) {
    if (-not (Test-Path -LiteralPath $dir)) { return $false }
    if (-not (Test-Writable $dir)) { return $false }
    $f = Get-ChildItem -LiteralPath $dir -Filter '*.tpl' -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($f) {
        if ($f.IsReadOnly) { return $false }
        return (Test-FileWritable $f.FullName)
    }
    return $true
}
function Clear-ReadOnlyTree([string]$dir) {
    try {
        Get-ChildItem -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                if (($_.Attributes -band [System.IO.FileAttributes]::ReadOnly) -ne 0) {
                    $_.Attributes = [System.IO.FileAttributes](([int]$_.Attributes) -band (-bnot [int][System.IO.FileAttributes]::ReadOnly))
                }
            } catch {}
        }
    } catch {}
}
# 客户机上常见“目录可写、现有 .tpl 只读 / ACL 拒绝写”。安装器自动修：
# 先清只读位；还不行再用系统自带 takeown/icacls 授权，全程不需要客户敲命令。
function Repair-InstallPermissions([string]$install) {
    $tpl = Join-Path $install (ConvertTo-PathSep (Get-EditionsCfg).paths.templates)
    Clear-ReadOnlyTree $install
    if ((Test-Path -LiteralPath $tpl) -and (Test-DirPatchable $tpl)) { return $true }
    $takeown = Join-Path $env:SystemRoot 'System32\takeown.exe'
    $icacls = Join-Path $env:SystemRoot 'System32\icacls.exe'
    $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    try { if (Test-Path -LiteralPath $takeown) { & $takeown /F $install /R /D Y 2>$null | Out-Null } } catch {}
    try { if (Test-Path -LiteralPath $icacls) { & $icacls $install /grant ($user + ':(OI)(CI)F') /T /C /Q 2>$null | Out-Null } } catch {}
    if ((Test-Path -LiteralPath $tpl) -and (Test-DirPatchable $tpl)) { return $true }
    return $false
}

function Test-IsAdmin {
    try {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        return (New-Object System.Security.Principal.WindowsPrincipal($id)).IsInRole(
            [System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

# ------------------------------------------------------------------ 数据加载
$script:EditionsCfg = $null
function Get-EditionsCfg {
    if (-not $script:EditionsCfg) {
        $p = Join-Path $DataRoot 'editions.json'
        if (-not (Test-Path -LiteralPath $p)) { throw "缺少数据文件：$p" }
        $script:EditionsCfg = Read-Json $p
    }
    return $script:EditionsCfg
}
function Get-Ed([string]$key) { return (Get-EditionsCfg).editions.$key }

$script:PatchGroups = $null
function Get-PatchGroups {
    if (-not $script:PatchGroups) {
        $list = New-Object System.Collections.ArrayList
        $dir = Join-Path $DataRoot 'patches'
        foreach ($f in (Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
            [void]$list.Add((Read-Json $f.FullName))
        }
        $script:PatchGroups = $list
    }
    return $script:PatchGroups
}
function Get-GroupWith([string]$key) {
    foreach ($g in (Get-PatchGroups)) { if ($g.$key) { return $g } }
    return $null
}
function Get-AllPatches {
    $out = New-Object System.Collections.ArrayList
    foreach ($g in (Get-PatchGroups)) { foreach ($p in $g.patches) { [void]$out.Add($p) } }
    return $out
}

function Get-PolicyBody([string]$lang, [string]$levelOverride = '') {
    $lv = $Level
    if ($levelOverride) { $lv = $levelOverride }
    $suffix = ''
    if ($lv -eq 'strong') { $suffix = '.strong' }
    $p = Join-Path (Join-Path $DataRoot 'policies') ("policy.$lang$suffix.txt")
    if (Test-Path -LiteralPath $p) {
        try { return ([System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8)).Trim() } catch {}
    }
    return $null
}

# ------------------------------------------------------------------ 政策块 / 文本
function Get-PolicyBlock([string]$text, [string]$open, [string]$close) {
    if (-not $text) { return $null }
    $a = $text.IndexOf($open)
    if ($a -lt 0) { return $null }
    $s = $a + $open.Length
    $b = $text.IndexOf($close, $s)
    if ($b -lt 0) { return $null }
    return $text.Substring($s, $b - $s)
}

function Set-PolicyBlock([string]$text, [string]$open, [string]$close, [string]$body, [string]$eol, [ref]$count) {
    $count.Value = 0
    $sb = New-Object System.Text.StringBuilder
    $cursor = 0
    while ($true) {
        $a = $text.IndexOf($open, $cursor)
        if ($a -lt 0) { break }
        $b = $text.IndexOf($close, $a + $open.Length)
        if ($b -lt 0) { break }
        $inner = $text.Substring($a + $open.Length, $b - $a - $open.Length)
        $isPlain = ($inner.Contains("`n") -or $inner.Contains("`r"))
        [void]$sb.Append($text.Substring($cursor, $a - $cursor))
        [void]$sb.Append($open)
        if ($isPlain) { [void]$sb.Append($eol + (ConvertTo-Nl $body $eol) + $eol) }
        else { [void]$sb.Append('\n' + (Get-JsonBody $body) + '\n') }
        [void]$sb.Append($close)
        $cursor = $b + $close.Length
        $count.Value++
    }
    [void]$sb.Append($text.Substring($cursor))
    return $sb.ToString()
}

function Test-SameBlock([string]$a, [string]$b) {
    $na = ((($a -replace "`r`n", "`n") -replace "`r", "`n").Trim())
    $nb = ((($b -replace "`r`n", "`n") -replace "`r", "`n").Trim())
    return ($na -ceq $nb)
}

function Resolve-PolicyLang([string]$blockText) {
    if ($PolicyLang -ne 'auto') { return $PolicyLang }
    if (-not $blockText) { return 'en' }
    $cjk = ([regex]::Matches($blockText, '[\u4e00-\u9fff]')).Count
    if ($cjk -gt ($blockText.Length * 0.02)) { return 'zh' }
    return 'en'
}

function Test-IsPatched([string]$text) {
    $g = Get-GroupWith 'content_policy'
    if (-not $g) { return $false }
    foreach ($m in $g.content_policy.is_patched_markers) { if ($text.Contains([string]$m)) { return $true } }
    return $false
}

# ------------------------------------------------------------------ 身份与探测
function Test-TemplatesDir([string]$dir) {
    if ([string]::IsNullOrWhiteSpace($dir)) { return $false }
    if (-not (Test-Path -LiteralPath $dir)) { return $false }
    $tpl = Join-Path $dir (ConvertTo-PathSep (Get-EditionsCfg).paths.templates)
    if (-not (Test-Path -LiteralPath $tpl)) { return $false }
    if (Test-Path -LiteralPath (Join-Path $tpl 'workbuddy-prompt.tpl')) { return $true }
    $any = Get-ChildItem -LiteralPath $tpl -Filter '*.tpl' -File -ErrorAction SilentlyContinue | Select-Object -First 1
    return ($null -ne $any)
}

# $true / $false / $null(未知)
function Test-Identity([string]$dir, $ed) {
    $pj = Join-Path $dir (ConvertTo-PathSep (Get-EditionsCfg).paths.productJson)
    $ident = $null
    if (Test-Path -LiteralPath $pj) { try { $ident = Read-Json $pj } catch { $ident = $null } }
    if ($ident) {
        $hits = 0
        foreach ($c in $ed.identity.checks) {
            $v = ''
            try { $v = [string]$ident.($c.field) } catch { $v = '' }
            if ($v -and ($v -match $c.regex)) { $hits++ }
        }
        if ($ed.identity.mode -eq 'any-of') { return ($hits -gt 0) }
        return ($hits -eq 0)
    }
    foreach ($fx in @($ed.identity.fallbackExe)) {
        if ($fx -and (Test-Path -LiteralPath (Join-Path $dir $fx))) { return $true }
    }
    return $null
}

function Get-ProductMeta([string]$dir) {
    $pj = Join-Path $dir (ConvertTo-PathSep (Get-EditionsCfg).paths.productJson)
    if (-not (Test-Path -LiteralPath $pj)) { return $null }
    try { return (Read-Json $pj) } catch { return $null }
}

# ------------------------------------------------------------------ 自动更新阻断
# 客户端 UpdateService.win32 在构造时会对 TEMP\workbuddy-update-<arch>[-app] 执行
# mkdirSync(recursive)。把该路径换成同名文件后，构造函数抛异常，更新服务起不来：
# 不检查、不下载、不提示、不安装。路径可按 -UpdateCachePath 覆盖（自测用）。
function Get-UpdateCacheInfo([string]$edKey, [string]$install, [string]$override) {
    if ($override) { return @{ cachePath = $override } }
    $app = ''
    if ($install) {
        try { $m = Get-ProductMeta $install; if ($m) { $app = [string]$m.applicationName } } catch {}
    }
    if (-not $app) { if ($edKey -eq 'intl') { $app = 'workbuddy-ai' } else { $app = 'WorkBuddy' } }
    $slug = ($app.ToLower() -replace '[^a-z0-9._-]+', '-')
    $slug = $slug.Trim('-')
    if (-not $slug) { $slug = 'workbuddy' }
    $arch = 'x64'
    if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { $arch = 'arm64' }
    $suffix = ''
    if ($slug -ne 'workbuddy') { $suffix = '-' + $slug }
    $tmp = [System.IO.Path]::GetTempPath()
    return @{ cachePath = (Join-Path $tmp ('workbuddy-update-' + $arch + $suffix)) }
}

function Stop-ClientForUpdate([string]$edKey) {
    $ed = Get-Ed $edKey
    $pat = [string]$ed.winProcNameFilter
    if (-not $pat) { return }
    $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $pat })
    foreach ($pr in $running) { try { $null = $pr.CloseMainWindow() } catch {} }
    if ($running.Count -gt 0) {
        Start-Sleep -Seconds 3
        foreach ($pr in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $pat })) {
            try { Stop-Process -Id $pr.Id -Force -ErrorAction Stop } catch {}
        }
        Start-Sleep -Seconds 1
    }
}

function Invoke-UpdateBlock([string]$edKey, [string]$install, [string]$override) {
    $ed = Get-Ed $edKey
    $info = Get-UpdateCacheInfo $edKey $install $override
    $cache = [string]$info.cachePath
    $stateDir = Get-PatchRoot $ed
    if (-not (Test-Path -LiteralPath $stateDir)) { New-Item -ItemType Directory -Force -Path $stateDir | Out-Null }
    $statePath = Join-Path $stateDir 'update-block.json'
    $prev = $null
    if (Test-Path -LiteralPath $statePath) { try { $prev = Read-Json $statePath } catch { $prev = $null } }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backup = ''
    $shared = Join-Path $env:LOCALAPPDATA '@genieworkbuddy-desktop-updater'
    $sharedBackup = ''
    $marker = "WorkBuddy update cache disabled by patch tool`r`nDelete this file or run 17 to restore.`r`n"

    if (-not $override) { Stop-ClientForUpdate $edKey }

    if (Test-Path -LiteralPath $cache) {
        $item = Get-Item -LiteralPath $cache -Force
        if ($item.PSIsContainer) {
            $backup = $cache + '.disabled-' + $stamp
            try { Move-Item -LiteralPath $cache -Destination $backup -Force -ErrorAction Stop }
            catch {
                Write-Log ('  更新缓存目录无法移动（客户端可能仍在运行）：' + $_.Exception.Message) 'Yellow'
                return $false
            }
        } elseif ((Get-Content -LiteralPath $cache -Raw -ErrorAction SilentlyContinue) -match 'update cache disabled by patch tool') {
            $backup = ''
        } else {
            $b = $cache + '.bak-' + $stamp
            try { Move-Item -LiteralPath $cache -Destination $b -Force -ErrorAction Stop; $backup = $b } catch {}
        }
    }
    # 重复阻断：缓存已经是标记文件时，必须沿用上一次记录的备份路径，
    # 否则状态会被空 backupPath 覆盖，恢复时找不到原目录。
    if (-not $backup -and $prev -and $prev.backupPath) {
        $pb = [string]$prev.backupPath
        if ($pb -and (Test-Path -LiteralPath $pb)) { $backup = $pb }
    }
    if (Test-Path -LiteralPath $shared) {
        $sharedBackup = $shared + '.disabled-' + $stamp
        try { Move-Item -LiteralPath $shared -Destination $sharedBackup -Force -ErrorAction Stop }
        catch { $sharedBackup = ''; Write-Log ('  共享更新缓存无法移动：' + $_.Exception.Message) 'Yellow' }
    } elseif ($prev -and $prev.sharedBackup -and (Test-Path -LiteralPath ([string]$prev.sharedBackup))) {
        $sharedBackup = [string]$prev.sharedBackup
        if ($prev.sharedPath) { $shared = [string]$prev.sharedPath }
    }
    try {
        [System.IO.File]::WriteAllText($cache, $marker, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        Write-Log ('  无法创建更新阻断标记：' + $_.Exception.Message) 'Red'
        return $false
    }
    $state = [ordered]@{
        edition = $edKey
        cachePath = $cache
        backupPath = $backup
        sharedPath = $shared
        sharedBackup = $sharedBackup
        createdAt = (Get-Date).ToString('s')
    }
    [System.IO.File]::WriteAllText($statePath, ($state | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding($false)))
    Write-Log ('  自动更新已阻断：' + $cache)
    return $true
}

function Invoke-UpdateUnblock([string]$edKey, [string]$install, [string]$override) {
    $ed = Get-Ed $edKey
    $info = Get-UpdateCacheInfo $edKey $install $override
    $cache = [string]$info.cachePath
    $statePath = Join-Path (Get-PatchRoot $ed) 'update-block.json'
    $st = $null
    if (Test-Path -LiteralPath $statePath) { try { $st = Read-Json $statePath } catch { $st = $null } }
    if ($st -and $st.cachePath) { $cache = [string]$st.cachePath }
    if (Test-Path -LiteralPath $cache -PathType Leaf) {
        $c = Get-Content -LiteralPath $cache -Raw -ErrorAction SilentlyContinue
        if ($c -match 'update cache disabled by patch tool') { Remove-Item -LiteralPath $cache -Force -ErrorAction SilentlyContinue }
    }
    if ($st -and $st.backupPath) {
        $bak = [string]$st.backupPath
        if ($bak -and (Test-Path -LiteralPath $bak)) {
            if (Test-Path -LiteralPath $cache) { Remove-Item -LiteralPath $cache -Force -ErrorAction SilentlyContinue }
            try { Move-Item -LiteralPath $bak -Destination $cache -Force -ErrorAction Stop }
            catch { Write-Log ('  更新缓存目录恢复失败：' + $_.Exception.Message) 'Yellow' }
        }
    }
    if ($st -and $st.sharedBackup) {
        $sb = [string]$st.sharedBackup
        if ($sb -and (Test-Path -LiteralPath $sb)) {
            if (Test-Path -LiteralPath $st.sharedPath) { Remove-Item -LiteralPath $st.sharedPath -Recurse -Force -ErrorAction SilentlyContinue }
            try { Move-Item -LiteralPath $sb -Destination ([string]$st.sharedPath) -Force -ErrorAction Stop } catch {}
        }
    }
    if (Test-Path -LiteralPath $statePath) { Remove-Item -LiteralPath $statePath -Force -ErrorAction SilentlyContinue }
    Write-Log ('  自动更新阻断已撤销：' + $cache)
    return $true
}

# ------------------------------------------------------------------ 备用前置入口
# 内置模型默认把请求发到 https://www.workbuddy.cn/v2（消费端网关，带输出内容审核）。
# 客户端自己把 https://copilot.tencent.com 列为 CN 的备用前置门
# （CN_SAAS_CORE_API_ALTERNATE_ORIGIN，同一套 token 受信），并支持用
# CODEBUDDY_BASE_URL 覆盖模型请求的 baseURL（settings.json 的 env 优先于进程环境变量）。
# 本开关只改模型请求的落点；鉴权、配置、上报、更新仍走默认门。
# 值必须带 /v2：客户端把它原样当 OpenAI baseURL，SDK 再拼 /chat/completions。
function Get-AltFrontDoorValue { return 'https://copilot.tencent.com/v2' }

function Get-AltFrontDoorState([string]$edKey) {
    $ed = Get-Ed $edKey
    $p = Join-Path (Get-PatchRoot $ed) 'altfrontdoor.json'
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    try { return (Read-Json $p) } catch { return $null }
}

# 改 settings.json 的 env.<name>；$value 为 $null 表示删除该键
function Set-SettingsEnvValue([string]$path, [string]$name, $value) {
    $obj = $null
    if (Test-Path -LiteralPath $path) { $obj = Read-Json $path }
    if ($null -eq $obj) { $obj = New-Object psobject }
    if (-not $obj.PSObject.Properties['env']) {
        $obj | Add-Member -NotePropertyName 'env' -NotePropertyValue (New-Object psobject) -Force
    }
    if ($null -eq $value) {
        if ($obj.env.PSObject.Properties[$name]) { $obj.env.PSObject.Properties.Remove($name) }
    } elseif ($obj.env.PSObject.Properties[$name]) {
        $obj.env.$name = [string]$value
    } else {
        $obj.env | Add-Member -NotePropertyName $name -NotePropertyValue ([string]$value) -Force
    }
    $parent = [System.IO.Path]::GetDirectoryName($path)
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    [System.IO.File]::WriteAllText($path, ($obj | ConvertTo-Json -Depth 20), (New-Object System.Text.UTF8Encoding($false)))
}

function Set-AltFrontDoor([string]$edKey, [bool]$enable) {
    $ed = Get-Ed $edKey
    $dd = Resolve-DataDir $ed
    if (-not (Test-Path -LiteralPath $dd)) { Write-Log ('  数据目录不存在，跳过：' + $dd) 'Yellow'; return $false }
    $settings = Join-Path $dd 'settings.json'
    $stateRoot = Get-PatchRoot $ed
    if (-not (Test-Path -LiteralPath $stateRoot)) { New-Item -ItemType Directory -Path $stateRoot -Force | Out-Null }
    $statePath = Join-Path $stateRoot 'altfrontdoor.json'

    if ($enable) {
        $st = Get-AltFrontDoorState $edKey
        $hasPrev = ($st -and $st.PSObject.Properties['previous'])
        if (-not $hasPrev) {
            # 第一次开：整文件备份 + 记录 CODEBUDDY_BASE_URL 原值（可能为空）
            if (Test-Path -LiteralPath $settings) {
                try { Copy-Item -LiteralPath $settings -Destination (Join-Path $stateRoot 'settings.json.altfrontdoor.bak') -Force } catch {}
            }
            $cur = $null
            try { $cur = (Read-Json $settings).env.CODEBUDDY_BASE_URL } catch {}
            $rec = [ordered]@{ previous = $cur; value = (Get-AltFrontDoorValue); setAt = (Get-Date).ToString('s') }
            [System.IO.File]::WriteAllText($statePath, ($rec | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding($false)))
        }
        Set-SettingsEnvValue $settings 'CODEBUDDY_BASE_URL' (Get-AltFrontDoorValue)
        $now = $null
        try { $now = (Read-Json $settings).env.CODEBUDDY_BASE_URL } catch {}
        if ([string]$now -ne (Get-AltFrontDoorValue)) {
            Write-Log ('  写入后回读不一致，未启用：' + $settings) 'Red'
            return $false
        }
        Write-Log ('  备用前置入口已启用：' + (Get-AltFrontDoorValue) + '（' + $ed.shortName + '）')
        return $true
    }

    # 关闭：没开过就静默返回；开过就按记录还原原值（没有记录则删键）
    $st = Get-AltFrontDoorState $edKey
    $hasState = ($st -and $st.PSObject.Properties['previous'])
    $cur = $null
    try { $cur = (Read-Json $settings).env.CODEBUDDY_BASE_URL } catch {}
    if (-not $hasState -and [string]$cur -ne (Get-AltFrontDoorValue)) { return $false }
    $restore = $null
    if ($hasState) { $restore = $st.previous }
    if (Test-Path -LiteralPath $settings) { Set-SettingsEnvValue $settings 'CODEBUDDY_BASE_URL' $restore }
    if (Test-Path -LiteralPath $statePath) { Remove-Item -LiteralPath $statePath -Force -ErrorAction SilentlyContinue }
    Write-Log ('  备用前置入口已关闭（' + $ed.shortName + '）')
    return $true
}

function Restart-ClientEdition([string]$edKey) {
    $ed = Get-Ed $edKey
    $install = ''
    try { if ($clientDirs.ContainsKey($edKey)) { $install = [string]$clientDirs[$edKey] } } catch {}
    if (-not $install) { Write-Log ('  未找到安装目录，请手动重启客户端（' + $ed.shortName + '）') 'Yellow'; return }
    $exe = Resolve-ExePath $install $ed (Get-ProductMeta $install)
    if (-not $exe) { Write-Log ('  未找到主程序，请手动重启客户端（' + $ed.shortName + '）') 'Yellow'; return }
    $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ed.winProcNameFilter })
    foreach ($pr in $running) { try { $null = $pr.CloseMainWindow() } catch {} }
    if ($running.Count -gt 0) {
        $deadline = (Get-Date).AddSeconds(8)
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 500
            if (@(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ed.winProcNameFilter }).Count -eq 0) { break }
        }
        foreach ($pr in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ed.winProcNameFilter })) {
            try { Stop-Process -Id $pr.Id -Force -ErrorAction Stop } catch {}
        }
        Start-Sleep -Seconds 1
    }
    try {
        $exeDir = [System.IO.Path]::GetDirectoryName($exe)
        $si = New-Object System.Diagnostics.ProcessStartInfo
        $si.FileName = $env:ComSpec
        $si.Arguments = '/c start "" /d "' + $exeDir + '" "' + $exe + '"'
        $si.UseShellExecute = $false
        $si.CreateNoWindow = $true
        [void][System.Diagnostics.Process]::Start($si)
        Write-Log ('  已重新启动（' + $ed.shortName + '）')
    } catch { Write-Log ('  启动失败：' + $_.Exception.Message) 'Yellow' }
}

function Find-ClientDirs {
    $cands = New-Object System.Collections.ArrayList
    $seen = @{}
    $add = {
        param($p)
        if (-not $p) { return }
        $p = ([string]$p).Trim()
        if (-not $p) { return }
        try { $p = $p.TrimEnd('\') } catch {}
        if ($p -notmatch '^[A-Za-z]:' -and $p -notmatch '^\\\\') { return }
        if (-not $seen.ContainsKey($p)) { $seen[$p] = $true; [void]$cands.Add($p) }
    }

    if ($InstallDir) { & $add $InstallDir }

    try {
        foreach ($pr in (Get-Process -ErrorAction SilentlyContinue)) {
            foreach ($k in @('cn', 'intl')) {
                if ($pr.Name -match (Get-Ed $k).winProcNameFilter) {
                    try {
                        if ($pr.Path) {
                            $d = [System.IO.Path]::GetDirectoryName($pr.Path)
                            & $add $d
                            & $add ([System.IO.Path]::GetDirectoryName($d))
                        }
                    } catch {}
                }
            }
        }
    } catch {}

    foreach ($key in @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')) {
        try {
            foreach ($it in (Get-ItemProperty $key -ErrorAction SilentlyContinue)) {
                if ($it.InstallLocation) { & $add (([string]$it.InstallLocation).Trim().Trim('"').TrimEnd('\')) }
                if ($it.UninstallString) {
                    $u = [string]$it.UninstallString
                    $m1 = [regex]::Match($u, '^\s*"([^"]+)"')
                    if ($m1.Success) { & $add ([System.IO.Path]::GetDirectoryName($m1.Groups[1].Value)) }
                    else {
                        $m2 = [regex]::Match($u, '^\s*([A-Za-z]:\\[^\s]+\.exe)')
                        if ($m2.Success) { & $add ([System.IO.Path]::GetDirectoryName($m2.Groups[1].Value)) }
                    }
                }
            }
        } catch {}
    }

    $names = @()
    foreach ($k in @('cn', 'intl')) { $names += (Get-Ed $k).winInstallDirNames }
    $names = $names | Sort-Object -Unique
    foreach ($base in @((Join-Path $env:LOCALAPPDATA 'Programs'), $env:LOCALAPPDATA, $env:APPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not $base -or -not (Test-Path -LiteralPath $base)) { continue }
        foreach ($n in $names) { & $add (Join-Path $base $n) }
        try {
            foreach ($sub in (Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue)) {
                & $add $sub.FullName
                foreach ($n in $names) { & $add (Join-Path $sub.FullName $n) }
            }
        } catch {}
    }
    foreach ($root in (Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^[A-Za-z]$' })) {
        foreach ($n in $names) { & $add (Join-Path $root.Root $n) }
    }

    try {
        $sh = New-Object -ComObject WScript.Shell
        foreach ($r in @(
            (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'),
            (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs'),
            (Join-Path $env:USERPROFILE 'Desktop'),
            (Join-Path $env:PUBLIC 'Desktop'))) {
            if (-not (Test-Path -LiteralPath $r)) { continue }
            foreach ($lnk in (Get-ChildItem -LiteralPath $r -Recurse -Filter '*.lnk' -ErrorAction SilentlyContinue)) {
                try {
                    $t = $sh.CreateShortcut($lnk.FullName).TargetPath
                    if ($t) {
                        $d = [System.IO.Path]::GetDirectoryName($t)
                        & $add $d
                        & $add ([System.IO.Path]::GetDirectoryName($d))
                    }
                } catch {}
            }
        }
    } catch {}

    $found = @{}
    foreach ($c in $cands) {
        if (-not (Test-TemplatesDir $c)) { continue }
        foreach ($k in @('cn', 'intl')) {
            if ($found.ContainsKey($k)) { continue }
            $taken = $false
            foreach ($k2 in $found.Keys) { if ($found[$k2] -eq $c) { $taken = $true } }
            if ($taken) { continue }
            if ((Test-Identity $c (Get-Ed $k)) -eq $true) { $found[$k] = $c }
        }
    }

    if ($found.Count -eq 0) {
        Write-Log '按常规位置未找到，开始递归搜索模板文件…'
        $roots = @($env:LOCALAPPDATA, (Join-Path $env:LOCALAPPDATA 'Programs'), $env:APPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:USERPROFILE) |
            Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique
        foreach ($root in $roots) {
            try {
                foreach ($h in (Get-ChildItem -LiteralPath $root -Recurse -Depth 5 -Filter 'workbuddy-prompt.tpl' -File -ErrorAction SilentlyContinue |
                                Where-Object { $_.FullName -notmatch '\\node_modules\\|\\.git\\|\\Temp\\' })) {
                    $d = $h.FullName
                    for ($i = 0; $i -lt 5; $i++) { $d = [System.IO.Path]::GetDirectoryName($d) }
                    if (-not (Test-TemplatesDir $d)) { continue }
                    foreach ($k in @('cn', 'intl')) {
                        if ($found.ContainsKey($k)) { continue }
                        if ((Test-Identity $d (Get-Ed $k)) -eq $true) { $found[$k] = $d }
                    }
                }
            } catch {}
            if ($found.Count -ge 2) { break }
        }
    }
    return $found
}

function Resolve-DataDir($ed) {
    # $DataDir 是脚本参数（-DataDir）。注意 PowerShell 变量名大小写不敏感，
    # 调用方千万不要再用 $dataDir 这个局部名，否则会把上一个版本的值当成"用户指定"。
    if ($DataDir) { return $DataDir }
    try {
        $ov = [Environment]::GetEnvironmentVariable($ed.dataDirEnvOverride, 'Process')
        if ($ov -and (Test-Path -LiteralPath $ov)) { return $ov }
    } catch {}
    return (Join-Path $env:USERPROFILE $ed.dataFolder)
}

function Resolve-ExePath([string]$install, $ed, $meta) {
    $names = @()
    if ($meta) { try { if ($meta.win32ExecutableName) { $names += [string]$meta.win32ExecutableName } } catch {} }
    $names += $ed.winExeNames
    foreach ($n in $names) {
        if (-not $n) { continue }
        $p = Join-Path $install $n
        if (Test-Path -LiteralPath $p) { return $p }
    }
    try {
        $c = Get-ChildItem -LiteralPath $install -Filter '*.exe' -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -match $ed.winProcNameFilter } | Select-Object -First 1
        if ($c) { return $c.FullName }
    } catch {}
    return $null
}

# ------------------------------------------------------------------ 工具族判断
function ConvertTo-NormPath([string]$p) {
    if (-not $p) { return '' }
    $t = $p.ToLower() -replace '\\', '/'
    if (-not $t.EndsWith('/')) { $t = $t + '/' }
    return $t
}

function Test-ForeignPath([string]$path, $ed) {
    $p = ConvertTo-NormPath $path
    foreach ($m in $ed.foreignPathMarkers) {
        if ($p.Contains(([string]$m).ToLower())) { return $true }
    }
    return $false
}

function Test-OwnedEnvValue([string]$value, $ed) {
    if (-not $value) { return $false }
    $v = ConvertTo-NormPath $value
    foreach ($root in (Get-EditionsCfg).family.patchRootNames) {
        if ($v.Contains((ConvertTo-NormPath $root))) { return $true }
    }
    foreach ($f in (Get-EditionsCfg).family.envFileNames) {
        if ($v.EndsWith(([string]$f).ToLower())) { return $true }
    }
    return $false
}

# ------------------------------------------------------------------ 快照与还原
function Save-ManifestLine([string]$line) {
    Add-Content -LiteralPath (Join-Path $script:SnapDir 'manifest.tsv') -Value $line -Encoding UTF8
}

function Backup-OneFile([string]$file) {
    if (-not $script:SnapDir) { return $null }
    if (-not (Test-Path -LiteralPath $file)) { return $null }
    $flat = ($file -replace '[:\\/]', '_')
    $dst = Join-Path $script:SnapDir $flat
    if (-not (Test-Path -LiteralPath $dst)) { Copy-Item -LiteralPath $file -Destination $dst -Force -ErrorAction Stop }
    $sha = ''
    try { $sha = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash } catch {}
    Save-ManifestLine ($file + "`t" + $dst + "`t" + $sha)
    # 备份下来的这份可能本身就带着改动（打补丁前文件已被别的工具注入过）。
    # 不标出来，还原时就拿不到原厂内容，用户会以为"还原没生效"。
    try { if (Test-PatchedFile $file) { Save-ManifestLine ('PREPATCHED' + "`t" + $file) } } catch {}
    return $dst
}

function Test-PatchedText([string]$text) {
    if (-not $text) { return $false }
    $g = Get-GroupWith 'content_policy'
    if ($g) { foreach ($m in $g.content_policy.is_patched_markers) { if ($text.Contains([string]$m)) { return $true } } }
    foreach ($p in (Get-AllPatches)) {
        $v = [string]$p.verify_re
        if (-not $v) { continue }
        try { if ([regex]::IsMatch($text, $v)) { return $true } } catch {}
    }
    return $false
}

function Test-PatchedFile([string]$path) {
    try {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        return (Test-PatchedText $strict.GetString($bytes))
    } catch { return $false }
}

# 把 -StateDir 归一成"补丁根目录"：允许传补丁根、state 目录、或某一份快照目录
function Resolve-StateRoot([string]$sd) {
    if (-not $sd) { return $null }
    if (Test-Path -LiteralPath (Join-Path $sd 'manifest.tsv')) { $sd = Split-Path -Parent $sd }
    if ((Split-Path -Leaf $sd) -eq 'state') { $sd = Split-Path -Parent $sd }
    if (Test-Path -LiteralPath (Join-Path $sd 'state')) { return $sd }
    return $null
}

# 还原前关掉客户端：文件被占用时复制会失败，而失败以前是被静默吞掉的
function Stop-ClientsForRestore {
    $n = 0
    foreach ($k in @('cn', 'intl')) {
        $ed = Get-Ed $k
        $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ed.winProcNameFilter })
        if ($running.Count -eq 0) { continue }
        Write-Log ('  关闭 ' + $ed.shortName + ' 进程 ' + $running.Count + ' 个（还原需要文件不被占用）')
        foreach ($pr in $running) { try { $null = $pr.CloseMainWindow() } catch {} }
        $deadline = (Get-Date).AddSeconds(8)
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 400
            if (@(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ed.winProcNameFilter }).Count -eq 0) { break }
        }
        foreach ($pr in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ed.winProcNameFilter })) {
            try { Stop-Process -Id $pr.Id -Force -ErrorAction Stop } catch {}
        }
        $n += $running.Count
        Start-Sleep -Milliseconds 500
    }
    return $n
}

# 扫一个安装根的补丁靶点：还有没有我们的痕迹（含"从未被我们写过"的文件）
function Get-ResidueInInstall([string]$install, $ed) {
    $bad = New-Object System.Collections.ArrayList
    $dd = $null
    try { $dd = Resolve-DataDir $ed } catch {}
    foreach ($p in (Get-AllPatches)) {
        if ($p.platforms -and ($p.platforms -notcontains $Platform)) { continue }
        $tgts = New-Object System.Collections.ArrayList
        if ($p.locate) { foreach ($f in (Resolve-BundleTargets $install $p)) { [void]$tgts.Add($f) } }
        else { foreach ($f in (Resolve-ScopeTargets $install $dd $ed $p.scope '' ([bool]$p.all_plugin_versions))) { [void]$tgts.Add($f) } }
        foreach ($f in $tgts) {
            if (-not (Test-Path -LiteralPath $f)) { continue }
            if (Test-PatchedFile $f) { [void]$bad.Add($f) }
        }
    }
    return $bad
}

# 取一个文件里的 <content_policy> 块（没有就返回空）
function Get-PolicyBlockOf([string]$file) {
    try {
        $t = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
        $pc = (Get-GroupWith 'content_policy').content_policy
        return (Get-PolicyBlock $t ([string]$pc.tag_open) ([string]$pc.tag_close))
    } catch { return $null }
}

# 修复"旧版本插件缓存里的残留"：
#   这些文件是本工具早期打的补丁，备份已不存在，所以没法用备份还原。
#   做法：找同插件的原厂副本，把它的 <content_policy> 块搬过来替换掉我们注入的块
#   （块级替换，文件其余内容逐字节不动）。
# 返回 $true 表示已修复。
function Repair-StalePluginCache([string]$file) {
    if ($file -notmatch '\\plugins\\cache\\') { return $false }
    $pc = (Get-GroupWith 'content_policy').content_policy
    $open = [string]$pc.tag_open
    $close = [string]$pc.tag_close

    # 找原厂块：同插件的其他（干净的）版本缓存 → 安装目录里的同名插件
    $stock = $null
    $cands = New-Object System.Collections.ArrayList
    if ($file -match '^(.*\\plugins\\cache\\[^\\]+\\[^\\]+)\\[^\\]+\\prompt\.tpl$') {
        $pluginDir = $Matches[1]
        foreach ($d in (Get-ChildItem -LiteralPath $pluginDir -Directory -ErrorAction SilentlyContinue)) {
            $cand = Join-Path $d.FullName 'prompt.tpl'
            if (Test-Path -LiteralPath $cand) { [void]$cands.Add($cand) }
        }
    }
    # 安装目录：welcomemode-code → welcomemode\code
    if ($file -match '\\plugins\\cache\\[^\\]+\\([^\\]+)\\') {
        $pluginName = $Matches[1]
        $rel = $pluginName -replace '-', '\'
        $roots = @()
        try { foreach ($k in @('cn', 'intl')) { $r = $clientDirs[$k]; if ($r) { $roots += (Join-Path $r 'resources\app.asar.unpacked\resources\plugins') } } } catch {}
        foreach ($rt in $roots) {
            if (-not (Test-Path -LiteralPath $rt)) { continue }
            foreach ($sub in @($rel, ($pluginName -replace '^(\w+?)-(.*)$', '$1\$2'))) {
                foreach ($mkt in @('workbuddy-builtin', 'codebuddy-plugins-official')) {
                    $p = Join-Path (Join-Path $rt $mkt) (Join-Path $sub 'prompt.tpl')
                    if (Test-Path -LiteralPath $p) { [void]$cands.Add($p) }
                }
            }
        }
    }
    foreach ($c in $cands) {
        if ($c -ceq $file) { continue }
        if (Test-PatchedFile $c) { continue }          # 带我们改动的不能当原厂
        $b = Get-PolicyBlockOf $c
        if ($b) { $stock = $b; break }
    }
    if (-not $stock) { return $false }

    try {
        $raw = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
        if (-not (Get-PolicyBlock $raw $open $close)) { return $false }
        $cnt = 0
        $new = Set-PolicyBlock $raw $open $close $stock (Get-FileEol $raw) ([ref]$cnt)
        if ($cnt -lt 1) { return $false }
        if ($new -ceq $raw) { return $false }
        Write-FileTracked $file $new
        $back = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
        if ($back -cne $new) { return $false }
        return (-not (Test-PatchedFile $file))
    } catch { return $false }
}

# 还有一类缓存"出生就带改动"：客户端会把安装目录里已被替换的 welcomemode 模板原样复制成
# <数据目录>\plugins\cache\...\prompt.tpl。它从没被我们拍过快照，所以没有任何原厂备份可还原，
# 但生成时与安装目录的模板逐字节相同 —— 安装目录还原干净后，直接镜像回去即可。
function Repair-PluginCacheFromInstall([string]$file) {
    if ($file -notmatch '\\plugins\\cache\\([^\\]+)\\([^\\]+)\\') { return $false }
    $market = $Matches[1]
    $pluginName = $Matches[2]

    # 缓存属于哪个版本，就用哪个版本的安装目录（两版可能同时装、版本不同）
    $keys = @()
    $ddRoot = $null
    if ($file -match '^(.*?)\\plugins\\cache\\') { $ddRoot = $Matches[1] }
    if ($ddRoot) {
        $n = ConvertTo-NormPath $ddRoot
        foreach ($k in @('cn', 'intl')) {
            $suffix = '/' + ([string](Get-Ed $k).dataFolder).ToLower() + '/'
            if ($n.EndsWith($suffix)) { $keys = @($k); break }
        }
    }
    if ($keys.Count -eq 0) { $keys = @('cn', 'intl') }

    # welcomemode-code → welcomemode\code（另一种命名 a-b-c → a\b-c 也试一下）
    $rels = @(($pluginName -replace '-', '\'))
    if ($pluginName -match '^(\w+?)-(.*)$') { $rels += ($Matches[1] + '\' + $Matches[2]) }

    foreach ($k in $keys) {
        $r = $null
        try { $r = $clientDirs[$k] } catch {}
        if (-not $r) { continue }
        $rt = Join-Path $r 'resources\app.asar.unpacked\resources\plugins'
        foreach ($rel in $rels) {
            $src = Join-Path (Join-Path $rt $market) (Join-Path $rel 'prompt.tpl')
            if (-not (Test-Path -LiteralPath $src)) { continue }
            if (Test-PatchedFile $src) { continue }   # 安装目录那份也没还原干净，不能当原厂
            try {
                Copy-Item -LiteralPath $src -Destination $file -Force -ErrorAction Stop
                if (-not (Test-PatchedFile $file)) { return $true }
            } catch {}
        }
    }
    return $false
}

# 删除"旧版本插件缓存"里残留了我们改动的整个版本目录（显式开关才用）
function Invoke-PurgeStaleCaches($eds, [string]$clientVer) {
    $removed = 0
    foreach ($k in $eds) {
        if (-not $clientDirs.ContainsKey($k)) { continue }
        $ed = Get-Ed $k
        $dd = Resolve-DataDir $ed
        $cacheRoot = Join-Path $dd 'plugins\cache'
        if (-not (Test-Path -LiteralPath $cacheRoot)) { continue }
        $dirs = @{}
        foreach ($f in (Get-ChildItem -LiteralPath $cacheRoot -Recurse -Filter 'prompt.tpl' -File -ErrorAction SilentlyContinue)) {
            if (-not (Test-StalePluginCache $f.FullName $clientVer)) { continue }
            if (-not (Test-PatchedFile $f.FullName)) { continue }
            $verDir = Split-Path -Parent $f.FullName
            if ($verDir) { $dirs[$verDir] = $true }
        }
        foreach ($d in $dirs.Keys) {
            try {
                Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction Stop
                $removed++
                Write-Log ('  已删除陈旧插件缓存目录：' + $d) 'Yellow'
            } catch {
                Write-Log ('  删除失败：' + $d + ' : ' + $_.Exception.Message) 'Red'
            }
        }
    }
    return $removed
}

function Invoke-RestoreAll {
    Write-Log '=== 开始还原 ==='
    if (-not $StateDir) {
        foreach ($k in @('cn', 'intl')) { try { $null = Invoke-UpdateUnblock $k '' '' } catch {} }
    }
    # 备用前置入口也是可选项：还原必须一并摘掉，否则"还原了但请求还走备用门"
    foreach ($k in @('cn', 'intl')) { try { $null = Set-AltFrontDoor $k $false } catch {} }
    $roots = New-Object System.Collections.ArrayList
    if ($StateDir) {
        # 指定了快照位置就只用这一处：便于隔离测试，或只还原某一个版本
        $r = Resolve-StateRoot $StateDir
        if ($r) { [void]$roots.Add(@{ root = $r; name = (Split-Path -Leaf $r) }) }
        else { Write-Log ('  -StateDir 指向的位置里没有 state 目录：' + $StateDir) 'Yellow' }
    } else {
        foreach ($rn in (Get-EditionsCfg).family.patchRootNames) {
            $r = Join-Path $env:LOCALAPPDATA $rn
            if (Test-Path -LiteralPath (Join-Path $r 'state')) { [void]$roots.Add(@{ root = $r; name = $rn }) }
        }
    }
    if ($roots.Count -eq 0) { Show-Message 'WorkBuddy 还原' '未找到任何备份记录，无需还原。'; return }

    $edOfRoot = @{}
    foreach ($k in @('cn', 'intl')) { $edOfRoot[(Get-Ed $k).patchRootName.ToLower()] = $k }

    $anyBackup = @{}          # 每个文件最新的一份备份
    $pristine = @{}           # 每个文件"看起来没被我们改过"的最新备份
    $deleteRestore = @{}      # 我们删掉过的文件 → 备份（只在仍缺失时放回）
    $writtenHashes = @{}      # 我们写出来过的内容指纹（黑名单）
    $prePatched = @{}         # 备份本身就带改动的文件（没有原厂备份可用，只能按内容摘）
    $installRoots = @{}       # 清单里记下的安装根（用于靶点级残留校验）
    $envNamesSeen = @{}       # 清单里记录过的环境变量名（客户端将来改名也能清掉）
    $runs = 0

    foreach ($r in $roots) {
        $rd = Join-Path $r.root 'state'
        if (-not (Test-Path -LiteralPath $rd)) { continue }
        $ek = $edOfRoot[$r.name.ToLower()]
        # 还原也要写日志：否则"仍有残留"这种关键结论只打到控制台，事后查不到
        try {
            if (-not (Test-Path -LiteralPath $r.root)) { New-Item -ItemType Directory -Path $r.root -Force | Out-Null }
            $script:LogFile = Join-Path $r.root 'install.log'
        } catch {}
        $ed = $null
        if ($ek) { $ed = Get-Ed $ek }
        foreach ($snap in (Get-ChildItem -LiteralPath $rd -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending)) {
            $man = Join-Path $snap.FullName 'manifest.tsv'
            if (-not (Test-Path -LiteralPath $man)) { continue }
            $runs++
            foreach ($line in (Get-Content -LiteralPath $man -Encoding UTF8)) {
                $parts = $line -split "`t"
                if ($parts.Count -lt 2) { continue }
                $key = $parts[0]
                if ($key -eq 'ENVFILE') {
                    if ($parts[1] -and (Test-Path -LiteralPath $parts[1])) {
                        try { Remove-Item -LiteralPath $parts[1] -Force -ErrorAction SilentlyContinue } catch {}
                    }
                    continue
                }
                if ($key -eq 'FILEDEL') {
                    # 我们删过的文件（运行时缓存）：只在它仍然缺失时放回，不覆盖客户端重新生成的版本
                    if ($parts.Count -ge 3 -and $parts[1] -and $parts[2] -and (Test-Path -LiteralPath $parts[2])) {
                        if (-not $deleteRestore.ContainsKey($parts[1])) { $deleteRestore[$parts[1]] = $parts[2] }
                    }
                    continue
                }
                if ($key -eq 'INSTALLROOT') {
                    if ($parts[1]) { $installRoots[$parts[1]] = $true }
                    continue
                }
                if ($key -eq 'PREPATCHED') {
                    if ($parts[1]) { $prePatched[$parts[1]] = $true }
                    continue
                }
                if ($key -eq 'WROTE') {
                    if ($parts.Count -ge 3 -and $parts[2]) { $writtenHashes[$parts[2].ToUpper()] = $true }
                    continue
                }
                if ($key -eq 'ENVNAME') {
                    # 变量名是打补丁时从 product.json 读出来的；内置列表将来可能对不上，
                    # 所以这里以清单记录为准，值属于本工具就一并清掉
                    if ($parts[1]) { $envNamesSeen[$parts[1]] = $true }
                    continue
                }
                if ($key -in @('ENVPREV', 'PROFILE')) { continue }
                $file = $parts[0]; $bak = $parts[1]
                $sha = ''
                if ($parts.Count -ge 3) { $sha = [string]$parts[2] }
                if (-not $file -or -not $bak) { continue }
                if ($ed -and (Test-ForeignPath $file $ed)) { continue }
                if (-not (Test-Path -LiteralPath $bak)) { continue }
                if (-not $anyBackup.ContainsKey($file)) { $anyBackup[$file] = $bak }
                if ($pristine.ContainsKey($file)) { continue }
                # 第一层：这份备份的内容，是不是我们写出来的？（对无标记产物最关键）
                if ($sha -and $writtenHashes.ContainsKey($sha.ToUpper())) { continue }
                # 第二/三层：内容含政策标记，或命中任何补丁的"已打"自检正则
                if (Test-PatchedFile $bak) { continue }
                $pristine[$file] = $bak
            }
        }
    }
    if ($runs -eq 0) { Show-Message 'WorkBuddy 还原' '快照目录里没有可用的清单，无需还原。'; return }

    $null = Stop-ClientsForRestore

    $restored = 0
    $failed = New-Object System.Collections.ArrayList
    foreach ($file in $anyBackup.Keys) {
        $src = $anyBackup[$file]
        if ($pristine.ContainsKey($file)) { $src = $pristine[$file] }
        $parent = [System.IO.Path]::GetDirectoryName($file)
        if ($parent -and -not (Test-Path -LiteralPath $parent)) { continue }
        try {
            Copy-Item -LiteralPath $src -Destination $file -Force -ErrorAction Stop
            $restored++
        } catch {
            [void]$failed.Add((Split-Path -Leaf $file) + ' : ' + $_.Exception.Message)
            Write-Log ('  [FAIL] 还原失败：' + $file + ' : ' + $_.Exception.Message) 'Red'
        }
    }

    $delRestored = 0
    foreach ($tgt in $deleteRestore.Keys) {
        if (Test-Path -LiteralPath $tgt) { continue }
        try {
            $parent = [System.IO.Path]::GetDirectoryName($tgt)
            if ($parent -and -not (Test-Path -LiteralPath $parent)) { continue }
            Copy-Item -LiteralPath $deleteRestore[$tgt] -Destination $tgt -Force -ErrorAction Stop
            $delRestored++
            Write-Log ('  已放回被删除的文件：' + $tgt)
        } catch { [void]$failed.Add((Split-Path -Leaf $tgt) + ' : ' + $_.Exception.Message) }
    }

    foreach ($k in @('cn', 'intl')) { Clear-OwnedEnv (Get-Ed $k) }
    foreach ($n in $envNamesSeen.Keys) {
        $v = [Environment]::GetEnvironmentVariable($n, 'User')
        if (-not $v) { continue }
        if (Test-OwnedEnvValue $v (Get-Ed 'cn') -or (Test-OwnedEnvValue $v (Get-Ed 'intl'))) {
            [Environment]::SetEnvironmentVariable($n, '', 'User')
            Write-Log ('  已清理环境变量（按清单记录）：' + $n)
        } else {
            Write-Log ('  保留 ' + $n + '（不属于本工具）：' + $v)
        }
    }

    # 还原校验：再扫一遍"还有没有我们的东西"，避免"还原了但没生效"无人察觉
    # 先把"陈旧插件缓存"就地修好（这些文件没有备份可还原，只能按内容修）
    # 注意：还原分支在常规探测之前就执行，$clientDirs 这时还是空的（修 install 目录的插件源要用它）
    if (-not $clientDirs) { try { $clientDirs = Find-ClientDirs } catch {} }
    $repaired = 0
    $mirrored = 0
    foreach ($k in @('cn', 'intl')) {
        # 不依赖安装目录探测：插件缓存只跟数据目录有关。客户端被卸载/移走、
        # 或探测失败时（$clientDirs 为空），<data>\plugins\cache 里的残留照样要修。
        $ddK = $null
        try { $ddK = Resolve-DataDir (Get-Ed $k) } catch {}
        if (-not $ddK) { continue }
        $cacheRootK = Join-Path $ddK 'plugins\cache'
        if (-not (Test-Path -LiteralPath $cacheRootK)) { continue }
        foreach ($f in (Get-ChildItem -LiteralPath $cacheRootK -Recurse -Filter 'prompt.tpl' -File -ErrorAction SilentlyContinue)) {
            if (-not (Test-PatchedFile $f.FullName)) { continue }
            if (Repair-StalePluginCache $f.FullName) {
                $repaired++
                Write-Log ('  已修复陈旧插件缓存：' + $f.FullName)
                continue
            }
            # 整段替换类补丁（契约模板）没有政策块，上面的块级修复无能为力；
            # 这类缓存多半是"出生就带改动"，只能从安装目录的原厂模板重建。
            if (Repair-PluginCacheFromInstall $f.FullName) {
                $mirrored++
                Write-Log ('  已按安装目录原厂模板重建插件缓存：' + $f.FullName)
            }
        }
    }
    if ($repaired -gt 0) { Write-Log ('  陈旧插件缓存修复：' + $repaired + ' 个') }
    if ($mirrored -gt 0) { Write-Log ('  插件缓存按原厂模板重建：' + $mirrored + ' 个') }

    $still = New-Object System.Collections.ArrayList
    foreach ($file in $anyBackup.Keys) {
        if (-not (Test-Path -LiteralPath $file)) { continue }
        try { if (Test-PatchedFile $file) { [void]$still.Add($file) } } catch {}
    }
    foreach ($ir in $installRoots.Keys) {
        if (-not (Test-Path -LiteralPath $ir)) { continue }
        $edGuess = Get-Ed 'cn'
        foreach ($k in @('cn', 'intl')) {
            $idr = $null
            try { $idr = Test-Identity $ir (Get-Ed $k) } catch {}
            if ($idr -eq $true) { $edGuess = Get-Ed $k; break }
        }
        foreach ($f in (Get-ResidueInInstall $ir $edGuess)) {
            if ($still -notcontains $f) { [void]$still.Add($f) }
        }
    }
    Write-Log ("还原完成，恢复文件数：" + $restored + "，放回删除的文件：" + $delRestored + "，失败：" + $failed.Count)

    $msg = "已恢复 " + $restored + " 个文件"
    if ($delRestored -gt 0) { $msg += "，放回 " + $delRestored + " 个被删除的文件" }
    if ($repaired -gt 0) { $msg += "，修复旧版本插件缓存 " + $repaired + " 个" }
    if ($mirrored -gt 0) { $msg += "，按原厂模板重建插件缓存 " + $mirrored + " 个" }
    $msg += "，并清理本工具写入的环境变量。"
    if ($still.Count -gt 0) {
        $msg += "`n`n仍检测到 " + $still.Count + " 个文件带着本工具的改动：`n"
        foreach ($x in ($still | Select-Object -First 6)) { $msg += ("  · " + $x + "`n") }
        $note = 0
        foreach ($x in $still) { if ($prePatched.ContainsKey($x)) { $note++ } }
        $cacheLeft = 0
        foreach ($x in $still) { if ($x -match '\\plugins\\cache\\') { $cacheLeft++ } }
        if ($cacheLeft -gt 0) {
            $msg += "`n其中 " + $cacheLeft + " 个是插件缓存。它们没有原厂备份（生成时安装目录里的模板就已被替换），`n而这次也没能从安装目录重建（安装目录缺失，或那里的模板同样带改动）。`n关闭客户端后重跑一次还原；仍不行就删掉该版本号目录，客户端下次启动会自动重建。"
        } elseif ($note -gt 0) {
            $msg += "`n其中 " + $note + " 个在备份时就已经带着改动（例如之前用过别的注入工具），`n所以没有原厂内容可用。请用补救命令按内容摘除： install.ps1 -DeepClean"
        } else {
            $msg += "`n若是" + '"插件缓存里的旧版本残留"' + "（路径含 plugins\cache），可用：`n  install.ps1 -PurgeStaleCaches（删除旧版本缓存目录，客户端需要时会重建）"
        }
        Write-Log '  还原校验：仍有文件带改动' 'Red'
        foreach ($x in $still) { Write-Log ('    - ' + $x) 'Red' }
    } elseif ($failed.Count -gt 0) {
        $msg += "`n`n有 " + $failed.Count + " 个文件还原失败（多为文件被占用或权限不足）。`n请关闭客户端后重跑一次 2-一键还原.bat。"
    } else {
        $msg += "`n`n校验通过：没有残留改动。"
        # 干净的时候也在控制台明说一句（否则用户只看到"还原完成"，不知道有没有校验过）
        Write-Log '  还原校验：没有残留改动' 'Green'
    }
    if ($failed.Count -gt 0) {
        Write-Log '  以下文件还原失败：' 'Red'
        foreach ($x in $failed) { Write-Log ('    - ' + $x) 'Red' }
    }
    $msg += "`n`n客户端已关闭，请手动重新打开。"
    Show-Message 'WorkBuddy 还原完成' $msg
}

# ------------------------------------------------------------------ 回读复查
# 同一文件可能被多条补丁先后写多次：pending 必须按路径去重，只留最后一次的预期内容，
# 否则会拿中间版本去比对，报出"写入未生效"的假失败（参考材料 06 的坑）。
function Write-FileTracked([string]$path, [string]$content) {
    # 客户机上 .tpl 可能带只读位；WriteAllText 对只读文件会抛“访问被拒绝”，先清掉。
    try {
        $fi = Get-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        if ($fi -and $fi.IsReadOnly) { $fi.IsReadOnly = $false }
    } catch {}
    [System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding($false)))
    for ($i = 0; $i -lt $script:Pending.Count; $i++) {
        if ($script:Pending[$i].Path -ceq $path) {
            $script:Pending[$i] = [pscustomobject]@{ Path = $path; Expect = $content }
            Record-Written $path
            return
        }
    }
    [void]$script:Pending.Add([pscustomobject]@{ Path = $path; Expect = $content })
    Record-Written $path
}

# 记账：把"我们写出来的内容"的指纹记进快照清单（WROTE 行）。
# 还原时据此判定"这份备份是不是我们打过的"。只靠政策文本标记不够用：
# 命令闸门注入、子句替换、文件保护中和的产物里根本没有政策标记，
# 只按标记判断会把"已被我们改过的备份"当成原厂备份，还原就还原不干净。
function Record-Written([string]$path) {
    if (-not $script:SnapDir) { return }
    try {
        $sha = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        Save-ManifestLine ('WROTE' + "`t" + $path + "`t" + $sha)
    } catch {}
}

function Invoke-Recheck($eds) {
    $bad = 0
    foreach ($p in $script:Pending) {
        try {
            $got = [System.IO.File]::ReadAllText($p.Path, [System.Text.Encoding]::UTF8)
            if ($got -cne $p.Expect) { Write-Log ('  [FAIL] 写入未生效（回读不一致）：' + $p.Path) 'Red'; $bad++ }
        } catch {
            Write-Log ('  [FAIL] 无法回读：' + $p.Path + ' : ' + $_.Exception.Message) 'Red'; $bad++
        }
    }
    # 闸门类补丁：只复查"本轮真正启用过"的，否则会把没开的补丁（例如 -WebFilter 未加）报成失败
    $dirs = $null
    foreach ($patch in (Get-AllPatches)) {
        if (-not $patch.verify_re) { continue }
        if ($patch.platforms -and ($patch.platforms -notcontains $Platform)) { continue }
        if (-not ($patch.op -eq 'insert_after' -or $patch.op -eq 'replace_literal')) { continue }
        if (-not $patch.locate) { continue }
        if ($script:EnabledPatchIds -notcontains $patch.id) { continue }
        if (-not $dirs) { $dirs = Find-ClientDirs }
        foreach ($k in $eds) {
            if (-not $dirs.ContainsKey($k)) { continue }
            foreach ($f in (Resolve-BundleTargets $dirs[$k] $patch)) {
                try {
                    $t = [System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8)
                    if (-not [regex]::IsMatch($t, [string]$patch.verify_re)) {
                        Write-Log ('  [FAIL] 补丁未生效 ' + $patch.id + ' → ' + (Split-Path -Leaf $f)) 'Red'; $bad++
                    }
                } catch {}
            }
        }
    }
    if ($bad -eq 0) { Write-Log '复查：全部就绪' 'Green' } else { Write-Log ("复查：$bad 处未生效") 'Red' }
    return $bad
}

# ------------------------------------------------------------------ 按内容摘除（DeepClean）
# 只在"没有干净备份可用"时用：不做备份对照，直接按补丁定义做逆操作。
# 只处理可逆类补丁（插入 / 整句替换）；政策块与文件保护块没有确定原文，不做。
function Invoke-UndoOnFile([string]$file, $patch, [ref]$info) {
    $op = [string]$patch.op
    $bytes = $null
    try {
        $bytes = [System.IO.File]::ReadAllBytes($file)
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        $raw = $strict.GetString($bytes)
    } catch { $info.Value = 'not-utf8'; return 0 }

    if ($op -eq 'insert_after') {
        # 逆操作：锚点之后紧跟的那串如果是我们插的，就精确摘掉
        $insert = [string]$patch.insert
        if (-not $insert) { $info.Value = 'skip'; return 0 }
        $m = [regex]::Match($raw, [string]$patch.anchor_re)
        if (-not $m.Success) { $info.Value = 'no-anchor'; return 0 }
        $b = $m.Index + $m.Length
        $need = $insert.Length
        if ($raw.Length -lt $b + $need) { $info.Value = 'skip'; return 0 }
        if ($raw.Substring($b, $need) -cne $insert) { $info.Value = 'clean'; return 0 }
        $new = $raw.Remove($b, $need)
        if ($new.Length -ne ($raw.Length - $need)) { $info.Value = 'length-mismatch'; return 0 }
    }
    elseif ($op -eq 'replace_many') {
        $new = $raw
        $n = 0
        foreach ($pair in $patch.pairs) {
            $cur = [string]$pair[1]; $orig = [string]$pair[0]
            if ($cur -and $new.Contains($cur)) {
                $n += ([regex]::Matches($new, [regex]::Escape($cur))).Count
                $new = $new.Replace($cur, $orig)
            }
        }
        if ($n -le 0) { $info.Value = 'clean'; return 0 }
    }
    elseif ($op -eq 'replace_literal') {
        $cur = [string]$patch.replace; $orig = [string]$patch.find
        if ($cur -and $raw.Contains($cur)) {
            $new = $raw.Replace($cur, $orig)
        } else { $info.Value = 'clean'; return 0 }
    }
    else { $info.Value = 'unsupported'; return 0 }

    if ($new -ceq $raw) { $info.Value = 'clean'; return 0 }
    if ($DryRun) { $info.Value = 'preview'; return 1 }
    $null = Backup-OneFile $file
    Write-FileTracked $file $new
    $info.Value = 'undone'
    return 1
}

function Invoke-DeepClean($eds) {
    Write-Log '=== DeepClean：按内容摘除可逆补丁（不依赖备份）==='
    $all = @(Get-AllPatches)
    [array]::Reverse($all)
    foreach ($k in $eds) {
        if (-not $clientDirs -or -not $clientDirs.ContainsKey($k)) { continue }
        $ed = Get-Ed $k
        $install = $clientDirs[$k]
        $patchRoot = Get-PatchRoot $ed
        if (-not (Test-Path -LiteralPath $patchRoot)) { New-Item -ItemType Directory -Path $patchRoot -Force | Out-Null }
        $script:LogFile = Join-Path $patchRoot $ed.logFileName
        Write-Log ('-- ' + $ed.displayName + ' : ' + $install)
        $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
        $script:SnapDir = Join-Path (Join-Path $patchRoot 'state') ($stamp + [string]$ed.snapshotSuffix)
        try {
            if (-not (Test-Path -LiteralPath $script:SnapDir)) { New-Item -ItemType Directory -Path $script:SnapDir -Force | Out-Null }
            [System.IO.File]::WriteAllText((Join-Path $script:SnapDir 'manifest.tsv'), '', (New-Object System.Text.UTF8Encoding($false)))
        } catch { $script:SnapDir = $null }
        foreach ($p in $all) {
            if ($p.platforms -and ($p.platforms -notcontains $Platform)) { continue }
            $op = [string]$p.op
            if ($op -notin @('insert_after', 'replace_many', 'replace_literal')) {
                if ($op -eq 'content_policy_block' -or $op -eq 'replace_tag_block') { continue }
            }
            $tgts = New-Object System.Collections.ArrayList
            if ($p.locate) { foreach ($f in (Resolve-BundleTargets $install $p)) { [void]$tgts.Add($f) } }
            else { foreach ($f in (Resolve-ScopeTargets $install (Resolve-DataDir $ed) $ed $p.scope '' ([bool]$p.all_plugin_versions))) { [void]$tgts.Add($f) } }
            $n = 0
            foreach ($f in $tgts) {
                $info = ''
                try { if ((Invoke-UndoOnFile $f $p ([ref]$info)) -gt 0) { $n++ } } catch {}
            }
            if ($n -gt 0) { Write-Log ('   ' + ([string]$p.id).PadRight(20) + ' 已摘除 ' + $n + ' 个文件') }
        }
    }
    $bad = Invoke-Recheck $eds
    Write-Log '=== DeepClean 完成 ==='
    return $bad
}

# ------------------------------------------------------------------ 补丁引擎
function Resolve-BundleTargets([string]$install, $patch) {
    $out = New-Object System.Collections.ArrayList
    if (-not $patch.locate) { return $out }
    $globRel = [string]$patch.locate.glob
    $anchor = [string]$patch.locate.anchor
    $idx = $globRel.IndexOf('/')
    if ($idx -lt 0) { return $out }
    $first = $globRel.Substring(0, $idx)
    $rest = $globRel.Substring($idx + 1)
    $searchRoot = Join-Path $install (ConvertTo-PathSep $first)
    if (-not (Test-Path -LiteralPath $searchRoot)) { return $out }
    $dirPart = [System.IO.Path]::GetDirectoryName((ConvertTo-PathSep $rest))
    $filePart = [System.IO.Path]::GetFileName($rest)
    $dirToScan = $searchRoot
    if ($dirPart) { $dirToScan = Join-Path $searchRoot $dirPart }
    if (-not (Test-Path -LiteralPath $dirToScan)) { return $out }
    foreach ($f in (Get-ChildItem -LiteralPath $dirToScan -Filter $filePart -File -ErrorAction SilentlyContinue)) {
        if ($anchor) {
            try { if (-not ([System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)).Contains($anchor)) { continue } } catch { continue }
        }
        [void]$out.Add($f.FullName)
    }
    return $out
}

# 判断要不要跳过这个插件缓存文件：
#   插件缓存按版本分目录：<data>\plugins\cache\<市场>\<插件>\<版本>-wb.<...>\prompt.tpl
#   客户端只读自己当前版本的目录；旧版本目录是升级后的遗留物。
#   ⚠ 客户端版本读不到时（首次运行、或文件被瞬时占用）一律跳过：
#     否则会把所有版本的缓存都打上补丁，正好制造出"没有备份可还原"的残留。
function Test-SkipPluginCache([string]$path, [string]$clientVer) {
    if ($path -notmatch '\\plugins\\cache\\') { return $false }
    if (-not $clientVer) { return $true }
    if ($path -match ('\\' + [regex]::Escape($clientVer) + '-wb\.')) { return $false }
    if ($path -match ('\\' + [regex]::Escape($clientVer) + '\\')) { return $false }
    return $true
}

# 严格版（用于"清理旧版本缓存"）：只有确认版本且不一致才算陈旧，版本未知时不动手
function Test-StalePluginCache([string]$path, [string]$clientVer) {
    if ($path -notmatch '\\plugins\\cache\\') { return $false }
    if (-not $clientVer) { return $false }
    if ($path -match ('\\' + [regex]::Escape($clientVer) + '-wb\.')) { return $false }
    if ($path -match ('\\' + [regex]::Escape($clientVer) + '\\')) { return $false }
    return $true
}

function Resolve-ScopeTargets([string]$install, [string]$dataDir, $ed, $scope, [string]$clientVer, [bool]$allPluginVersions = $false) {
    $cfg = Get-EditionsCfg
    $out = New-Object System.Collections.ArrayList
    $seen = @{}
    $scopes = @()
    if ($scope -is [string]) { $scopes = @($scope) } else { $scopes = $scope }

    if ($scopes -contains 'templates') {
        $tplDir = Join-Path $install (ConvertTo-PathSep $cfg.paths.templates)
        if (Test-Path -LiteralPath $tplDir) {
            foreach ($f in (Get-ChildItem -LiteralPath $tplDir -Recurse -File -ErrorAction SilentlyContinue |
                            Where-Object { $_.Extension -eq '.tpl' -or $_.Extension -eq '.md' })) {
                if (-not $seen.ContainsKey($f.FullName)) { $seen[$f.FullName] = $true; [void]$out.Add($f.FullName) }
            }
        }
        $plugDirs = @((Join-Path $install (ConvertTo-PathSep $cfg.paths.plugins)))
        if ($dataDir) { foreach ($sub in @('plugins\cache', 'plugins\marketplaces', 'plugins\data')) { $plugDirs += (Join-Path $dataDir $sub) } }
        foreach ($pd in $plugDirs) {
            if (-not (Test-Path -LiteralPath $pd)) { continue }
            foreach ($f in (Get-ChildItem -LiteralPath $pd -Recurse -Filter 'prompt.tpl' -File -ErrorAction SilentlyContinue)) {
                if ($f.FullName -notmatch 'welcomemode') { continue }
                # 只处理当前客户端版本的插件缓存：旧版本缓存客户端不读，
                # 打上去只会在以后变成"没有备份可还原"的残留
                if (-not $allPluginVersions -and (Test-SkipPluginCache $f.FullName $clientVer)) { $script:StaleSkipped++; continue }
                if (-not $seen.ContainsKey($f.FullName)) { $seen[$f.FullName] = $true; [void]$out.Add($f.FullName) }
            }
        }
    }
    if ($scopes -contains 'productJson') {
        $pj = Join-Path $install (ConvertTo-PathSep $cfg.paths.productJson)
        if (Test-Path -LiteralPath $pj) { if (-not $seen.ContainsKey($pj)) { $seen[$pj] = $true; [void]$out.Add($pj) } }
    }
    return $out
}

function Invoke-PatchOnFile([string]$file, $patch, $ed, [ref]$info) {
    # 字节安全闸门：原文件必须是合法 UTF-8。
    # 用替换回退的编码去读-改-写会把非法字节变成 U+FFFD（一个字节变三个），文件就毁了；
    # 客户端 JS bundle 里将来可能夹带非法字节，所以宁可拒绝也不冒险。
    $bytes = [System.IO.File]::ReadAllBytes($file)
    $strict = New-Object System.Text.UTF8Encoding($false, $true)
    try { $raw = $strict.GetString($bytes) } catch { $info.Value = 'not-utf8'; return 0 }

    $eol = Get-FileEol $raw
    $isJson = $file.ToLower().EndsWith('.json')
    $new = $raw
    $n = 0
    if ($patch.only_if_contains -and $patch.op -ne 'replace_file' -and -not $raw.Contains([string]$patch.only_if_contains)) { $info.Value = 'no-match'; return 0 }

    switch ([string]$patch.op) {
        'insert_after' {
            if ($patch.idempotent_at_insert) {
                $m0 = [regex]::Match($raw, [string]$patch.anchor_re)
                if ($m0.Success) {
                    $ins = $m0.Index + $m0.Length
                    $len = [Math]::Min(([string]$patch.insert).Length, $raw.Length - $ins)
                    if ($len -gt 0 -and $raw.Substring($ins, $len) -eq [string]$patch.insert) { $info.Value = 'uptodate'; return 0 }
                }
            }
            $rx = New-Object System.Text.RegularExpressions.Regex([string]$patch.anchor_re)
            $sb = New-Object System.Text.StringBuilder
            $pos = 0
            foreach ($m in $rx.Matches($raw)) {
                if ($m.Index -lt $pos) { continue }
                [void]$sb.Append($raw.Substring($pos, $m.Index - $pos + $m.Length))
                [void]$sb.Append([string]$patch.insert)
                $pos = $m.Index + $m.Length
                $n++
            }
            [void]$sb.Append($raw.Substring($pos))
            $new = $sb.ToString()
        }
        'replace_literal' {
            $find = [string]$patch.find
            if ($raw.Contains($find)) {
                $n = ([regex]::Matches($raw, [regex]::Escape($find))).Count
                $new = $raw.Replace($find, [string]$patch.replace)
            } else { $info.Value = 'uptodate'; return 0 }
        }
        'replace_many' {
            foreach ($pair in $patch.pairs) {
                $old = [string]$pair[0]
                if ($new.Contains($old)) {
                    $n += ([regex]::Matches($new, [regex]::Escape($old))).Count
                    $new = $new.Replace($old, [string]$pair[1])
                }
            }
        }
        'replace_file' {
            # 整文件替换：用于 resources/templates/*.tpl 这类“整份系统提示”载体。
            # only_if_contains 先筛靶点，避免把 system-reminder 之类的小模板也换掉。
            $rt = [System.IO.File]::ReadAllText((Join-Path (Join-Path $DataRoot 'patches') ([string]$patch.replace_with_file)), [System.Text.Encoding]::UTF8)
            $rt = (ConvertTo-Nl $rt.TrimEnd("`r", "`n") $eol) + $eol
            if (($raw.TrimEnd("`r", "`n")) -ceq ($rt.TrimEnd("`r", "`n"))) { $info.Value = 'uptodate'; return 0 }
            if ($patch.only_if_contains -and -not $raw.Contains([string]$patch.only_if_contains)) { $info.Value = 'no-match'; return 0 }
            $new = $rt
            $n = 1
        }
        'replace_tag_block' {
            $rt = [System.IO.File]::ReadAllText((Join-Path (Join-Path $DataRoot 'patches') ([string]$patch.replace_with_file)), [System.Text.Encoding]::UTF8)
            $rt = $rt.TrimEnd("`r", "`n")
            $safety = [string]$patch.safety_valve
            foreach ($tag in $patch.tags) {
                $op = '<' + $tag + '>'
                $cl = '</' + $tag + '>'
                $from = 0
                while ($true) {
                    $i = $new.IndexOf($op, $from)
                    if ($i -lt 0) { break }
                    $j = $new.IndexOf($cl, $i)
                    if ($j -lt 0) { break }
                    $body = $new.Substring($i, $j - $i)
                    # 安全阀：含"最高优先级覆盖"说明那是我们自己的覆盖层，不能替换。
                    # JSON 里块是转义形态，中文可能被写成 \uXXXX，所以再拿解码后的文本判一次。
                    $hitSafety = $false
                    if ($safety) {
                        if ($body.Contains($safety)) { $hitSafety = $true }
                        elseif (-not ($body.Contains("`n") -or $body.Contains("`r"))) {
                            try { if ((('"' + $body + '"') | ConvertFrom-Json).Contains($safety)) { $hitSafety = $true } } catch {}
                        }
                    }
                    if ($hitSafety) { $from = $j + $cl.Length; continue }
                    $isPlain = ($body.Contains("`n") -or $body.Contains("`r"))
                    if ($isJson -and -not $isPlain) { $repl = Get-JsonBody (ConvertTo-Nl $rt "`n") }
                    else { $repl = ConvertTo-Nl $rt $eol }
                    $new = $new.Substring(0, $i) + $repl + $new.Substring($j + $cl.Length)
                    $n++
                    $from = $i + $repl.Length
                }
            }
        }
        'content_policy_block' {
            $g = Get-GroupWith 'content_policy'
            $cpc = $g.content_policy
            $blk = Get-PolicyBlock $new ([string]$cpc.tag_open) ([string]$cpc.tag_close)
            if (-not $blk) { $info.Value = 'no-policy'; return 0 }
            $body = Get-PolicyBody (Resolve-PolicyLang $blk)
            if (-not $body) { $info.Value = 'no-policy-file'; return 0 }
            # JSON 里的块是转义形态（字面 \n），必须解码后再比，否则永远不相等 → 每次都重写
            $isPlainBlk = ($blk.Contains("`n") -or $blk.Contains("`r"))
            $cmp = $blk
            if ($isPlainBlk -eq $false) {
                try { $cmp = ('"' + $blk + '"') | ConvertFrom-Json } catch { $cmp = $blk }
            }
            if (Test-SameBlock $cmp $body) { $info.Value = 'uptodate'; return 0 }
            $c = 0
            $new = Set-PolicyBlock $new ([string]$cpc.tag_open) ([string]$cpc.tag_close) $body $eol ([ref]$c)
            $n = $c
        }
        'replace_prompt_template' {
            # 整段替换某个 prompt 的 template（JSON 字符串），文件其余字节不动。
            # 新模板 = patches/<template_file> 的正文 + 末尾 <content_policy> 块（正文按档位/语言运行时注入）
            $pname = [string]$patch.prompt_name
            $tfile = Join-Path (Join-Path $DataRoot 'patches') ([string]$patch.template_file)
            if (-not (Test-Path -LiteralPath $tfile)) { $info.Value = 'no-template'; return 0 }
            $ni = $new.IndexOf('"name": "' + $pname + '"', [System.StringComparison]::Ordinal)
            if ($ni -lt 0) { $ni = $new.IndexOf('"name":"' + $pname + '"', [System.StringComparison]::Ordinal) }
            if ($ni -lt 0) { $info.Value = 'no-prompt'; return 0 }
            $ti = $new.IndexOf('"template"', $ni, [System.StringComparison]::Ordinal)
            if ($ti -lt 0) { $info.Value = 'no-template-key'; return 0 }
            $q1 = $new.IndexOf('"', $new.IndexOf(':', $ti) + 1)
            if ($q1 -lt 0) { $info.Value = 'no-template-value'; return 0 }
            $q2 = -1; $k = $q1 + 1
            while ($k -lt $new.Length) {
                if ($new[$k] -eq '\') { $k += 2; continue }
                if ($new[$k] -eq '"') { $q2 = $k; break }
                $k++
            }
            if ($q2 -lt 0) { $info.Value = 'unterminated'; return 0 }
            $marker = [string]$patch.marker
            if (-not $marker) { $marker = 'You are a local assistant running on the user''s own computer' }
            if ($new.Substring($q1 + 1, $q2 - $q1 - 1).Contains($marker)) { $info.Value = 'uptodate'; return 0 }
            $pc = (Get-GroupWith 'content_policy').content_policy
            $open = [string]$pc.tag_open; $close = [string]$pc.tag_close
            $lang = 'zh'
            if ([string]$ed.key -eq 'intl') { $lang = 'en' }
            $pb = Get-PolicyBody $lang
            if (-not $pb) { $pb = Get-PolicyBody 'zh' }
            $tplText = ([System.IO.File]::ReadAllText($tfile, [System.Text.Encoding]::UTF8)).TrimEnd()
            if ($patch.omit_policy -eq $true) {
                $full = $tplText + "`n"
            } else {
                $full = $tplText + "`n`n" + $open + "`n" + $pb + "`n" + $close + "`n"
            }
            $esc = $full.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t')
            $new = $new.Substring(0, $q1 + 1) + $esc + $new.Substring($q2)
            $n = 1
        }
        default { $info.Value = 'unknown-op'; return 0 }
    }

    if ($n -le 0) { if (-not $info.Value) { $info.Value = 'no-match' }; return 0 }
    # 不变量自检：插入/替换的字符数差必须精确对上，对不上说明正则行为超出预期 → 放弃
    if ($patch.op -eq 'insert_after' -or $patch.op -eq 'replace_literal') {
        $delta = 0
        if ($patch.op -eq 'insert_after') { $delta = $n * ([string]$patch.insert).Length }
        else { $delta = $n * (([string]$patch.replace).Length - ([string]$patch.find).Length) }
        if (($new.Length - $raw.Length) -ne $delta) { $info.Value = 'length-mismatch'; return 0 }
    }
    if ($isJson) {
        try { $null = ($new | ConvertFrom-Json) } catch { $info.Value = 'json-invalid'; return 0 }
    }
    if ($DryRun) { $info.Value = 'preview'; return $n }
    $null = Backup-OneFile $file
    try {
        Write-FileTracked $file $new
    } catch [System.UnauthorizedAccessException] {
        $info.Value = 'access-denied'; return 0
    } catch [System.IO.IOException] {
        $info.Value = 'file-locked'; return 0
    }
    $info.Value = 'patched'
    return $n
}

# ------------------------------------------------------------------ 环境变量通道
function Get-PatchRoot($ed) {
    if ($StateDir) { return (Split-Path -Parent $StateDir) }
    return (Join-Path $env:LOCALAPPDATA $ed.patchRootName)
}
function Get-EnvBackupPath($ed) { return (Join-Path (Get-PatchRoot $ed) 'env-backup.json') }

function Save-EnvBackup($ed, [string]$name, [string]$value) {
    $p = Get-EnvBackupPath $ed
    $obj = $null
    if (Test-Path -LiteralPath $p) { try { $obj = Read-Json $p } catch { $obj = $null } }
    if (-not $obj) { $obj = New-Object psobject }
    if (-not ($obj.PSObject.Properties.Name -contains $name)) {
        $root = Get-PatchRoot $ed
        if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null }
        Add-Member -InputObject $obj -MemberType NoteProperty -Name $name -Value $value
        [System.IO.File]::WriteAllText($p, ($obj | ConvertTo-Json -Depth 3), (New-Object System.Text.UTF8Encoding($false)))
    }
}

function Set-EnvOwned($ed, [string]$name, [string]$envFile) {
    $cur = [Environment]::GetEnvironmentVariable($name, 'User')
    if ($cur -and -not (Test-OwnedEnvValue $cur $ed)) {
        Write-Log ('  注意：' + $name + ' 当前指向别的工具（' + $cur + '），原值已记录，还原时不动它')
    }
    Save-EnvBackup $ed $name $cur
    foreach ($n in (Get-EditionsCfg).envNames) {
        if ($n -eq $name) { continue }
        $v = [Environment]::GetEnvironmentVariable($n, 'User')
        if ($v -and (Test-OwnedEnvValue $v $ed)) {
            [Environment]::SetEnvironmentVariable($n, '', 'User')
            Write-Log ('  清理旧别名：' + $n)
        }
    }
    [Environment]::SetEnvironmentVariable($name, $envFile, 'User')
    # 记进清单：变量名是从客户端 product.json 读出来的，客户端将来改名时
    # 内置列表会对不上，还原要能按清单记录来清
    Save-ManifestLine ('ENVFILE' + "`t" + $envFile)
    Save-ManifestLine ('ENVNAME' + "`t" + $name)
}

function Clear-OwnedEnv($ed) {
    $envFile = Join-Path (Get-PatchRoot $ed) $ed.envFileName
    $bp = Get-EnvBackupPath $ed
    foreach ($n in (Get-EditionsCfg).envNames) {
        $v = [Environment]::GetEnvironmentVariable($n, 'User')
        if (-not $v) { continue }
        $isOurs = $false
        if ((ConvertTo-NormPath $v) -eq (ConvertTo-NormPath $envFile)) { $isOurs = $true }
        if (Test-OwnedEnvValue $v $ed) { $isOurs = $true }
        if (-not $isOurs) { Write-Log ('  保留 ' + $n + '（不属于本工具）：' + $v); continue }
        $done = $false
        if (Test-Path -LiteralPath $bp) {
            try {
                $b = Read-Json $bp
                if ($b.PSObject.Properties.Name -contains $n) {
                    $orig = [string]$b.$n
                    [Environment]::SetEnvironmentVariable($n, $orig, 'User')
                    if ($orig) { Write-Log ('  已恢复原值：' + $n + ' = ' + $orig) } else { Write-Log ('  已清除：' + $n) }
                    $done = $true
                }
            } catch {}
        }
        if (-not $done) { [Environment]::SetEnvironmentVariable($n, '', 'User'); Write-Log ('  已清除：' + $n) }
    }
    if (Test-Path -LiteralPath $envFile) { Remove-Item -LiteralPath $envFile -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $bp) { Remove-Item -LiteralPath $bp -Force -ErrorAction SilentlyContinue }
}

function Build-EnvConfig($ed, [string]$install, [string]$dataDir, [ref]$promptCount) {
    $promptCount.Value = 0
    # 数据源优先级：客户端实时缓存（最准）→ 安装目录 cli/product.json（缓存被我们清掉或客户端从没启动过时用）。
    # 用 product.json 兜底很关键：清缓存与建覆盖配置在同一次运行里发生，否则第二次跑就会"找不到实时配置"而跳过通道 A。
    # 只读 product.json 不算动 bundle，所以 mac 的 A 通道也能用这条兜底。
    $src = $null
    $srcKind = ''
    $cacheDir = Join-Path $dataDir 'cache'
    if (Test-Path -LiteralPath $cacheDir) {
        # 5.9.2 起国内版把运行时缓存改成"每个进程一份"：
        # acc-product-config-v3-<pid>-<guid>.json，固定文件名可能根本不存在。
        # 按修改时间取最新的一份，作为通道 A 的数据源。
        $cand = @(Get-ChildItem -LiteralPath $cacheDir -File -Filter 'acc-product-config*.json' -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1)
        if ($cand.Count -gt 0) { $src = $cand[0].FullName; $srcKind = 'cache' }
    }
    if (-not $src) {
        $pj = Join-Path $install (ConvertTo-PathSep (Get-EditionsCfg).paths.productJson)
        if (Test-Path -LiteralPath $pj) { $src = $pj; $srcKind = 'product.json' }
    }
    if (-not $src) { return @{ ok = $false; reason = 'no-source' } }
    $live = $null
    try { $live = Read-Json $src } catch { return @{ ok = $false; reason = 'unreadable-source' } }
    $pathEnv = [string]$live.productConfigPathEnv
    if (-not $pathEnv) { $pathEnv = 'ACC_PRODUCT_CONFIG_PATH' }

    $cpc = (Get-GroupWith 'content_policy').content_policy
    $out = New-Object System.Collections.ArrayList
    foreach ($p in $live.prompts) {
        $t = [string]$p.template
        if (-not $t) { continue }
        if (-not $t.Contains([string]$cpc.tag_open)) { continue }
        $blk = Get-PolicyBlock $t ([string]$cpc.tag_open) ([string]$cpc.tag_close)
        if (-not $blk) { continue }
        $body = Get-PolicyBody (Resolve-PolicyLang $blk)
        if (-not $body) { continue }
        if (Test-SameBlock $blk $body) { continue }
        $c = 0
        $newT = Set-PolicyBlock $t ([string]$cpc.tag_open) ([string]$cpc.tag_close) $body "`n" ([ref]$c)
        if ($c -lt 1) { continue }
        [void]$out.Add([ordered]@{ name = [string]$p.name; template = $newT })
        $promptCount.Value++
    }
    if ($promptCount.Value -eq 0) { return @{ ok = $false; reason = 'nothing-to-patch'; pathEnv = $pathEnv; srcKind = $srcKind } }
    return @{ ok = $true; json = (@{ prompts = @($out) } | ConvertTo-Json -Depth 6 -Compress); pathEnv = $pathEnv; srcKind = $srcKind; src = $src }
}

# ------------------------------------------------------------------ 守护任务
function Test-TaskExists([string]$name) {
    try { $null = & schtasks /Query /TN $name 2>$null; return ($LASTEXITCODE -eq 0) } catch { return $false }
}

function Test-TaskWellFormed([string]$name) {
    $xml = ''
    try { $xml = (& schtasks /Query /TN $name /XML 2>$null | Out-String) } catch { return @{ ok = $false; why = '查询失败' } }
    $m = [regex]::Match($xml, '<Command>(.*?)</Command>', 'Singleline')
    if (-not $m.Success) { return @{ ok = $false; why = '任务 XML 里找不到 <Command>' } }
    $cmd = $m.Groups[1].Value.Trim()
    $arg = ''
    $ma = [regex]::Match($xml, '<Arguments>(.*?)</Arguments>', 'Singleline')
    if ($ma.Success) { $arg = $ma.Groups[1].Value.Trim() }
    $sn = 'install.ps1'
    if ($arg -and $arg.ToLower().Contains($sn)) { return @{ ok = $true; why = '' } }
    if ((-not $arg) -and $cmd.ToLower().Contains($sn)) {
        return @{ ok = $false; why = '整条命令行被塞进 <Command>、没有 <Arguments> → Windows 会当成一个带空格的程序名，任务跑不起来' }
    }
    if ((-not $arg) -and $cmd.ToLower().EndsWith('.exe')) { return @{ ok = $false; why = '只有解释器、没有脚本参数 → 任务起来也不会做任何事' } }
    return @{ ok = $false; why = '命令行里没找到脚本名' }
}

function Get-TaskLastResult([string]$name) {
    $out = ''
    try { $out = (& schtasks /Query /TN $name /V /FO LIST 2>$null | Out-String) } catch { return $null }
    foreach ($ln in ($out -split "`r?`n")) {
        if ($ln -match '^\s*(Last Result|上次运行结果)\s*[:：]\s*(-?\d+)') { return [int]$Matches[2] }
    }
    return $null
}

function Install-GuardTasks($eds) {
    $psExe = 'powershell.exe'
    try { $p = (Get-Process -Id $PID).Path; if ($p) { $psExe = $p } } catch {}
    $tr = '"' + $psExe + '" -NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '" -Silent -NoRestart -NoDialog'
    $isAdmin = Test-IsAdmin
    $results = New-Object System.Collections.ArrayList
    foreach ($k in $eds) {
        $ed = Get-Ed $k
        foreach ($t in @(
            @{ n = ($ed.guardTaskTag + '_Hourly'); s = @('/SC', 'MINUTE', '/MO', '30'); adm = $false },
            @{ n = ($ed.guardTaskTag + '_Logon');  s = @('/SC', 'ONLOGON');              adm = $true })) {
            $argv = @('/Create', '/TN', $t.n) + $t.s + @('/TR', $tr, '/F')
            if ($t.adm -and $isAdmin) { $argv += @('/RL', 'HIGHEST') }
            try { $null = & schtasks @argv 2>&1 } catch {}
            $exists = Test-TaskExists $t.n
            $wf = @{ ok = $false; why = '未创建' }
            $lr = $null
            if ($exists) { $wf = Test-TaskWellFormed $t.n; $lr = Get-TaskLastResult $t.n }
            [void]$results.Add(@{ name = $t.n; exists = $exists; wellformed = $wf.ok; why = $wf.why; lastResult = $lr })
        }
    }
    return $results
}

function Remove-GuardTasks($eds) {
    $removed = 0; $failed = New-Object System.Collections.ArrayList
    foreach ($k in $eds) {
        $ed = Get-Ed $k
        foreach ($n in @(($ed.guardTaskTag + '_Hourly'), ($ed.guardTaskTag + '_Logon'))) {
            if (-not (Test-TaskExists $n)) { continue }
            try { $null = & schtasks /Delete /TN $n /F 2>&1 } catch {}
            if (Test-TaskExists $n) { [void]$failed.Add($n) } else { $removed++ }
        }
    }
    return @{ removed = $removed; failed = $failed }
}

# ------------------------------------------------------------------ 云记忆
function Get-CloudMemCfg { return (Get-GroupWith 'our_mark') }

function Test-IsOurCloudMem([string]$text) {
    $c = Get-CloudMemCfg
    if ($text.Contains([string]$c.our_mark)) { return $true }
    foreach ($m in $c.other_tool_marks) { if ($text.Contains([string]$m)) { return $true } }
    $m = [regex]::Match($text, '(?m)^\s*>\s*Version:\s*(\d+)\s*$')
    if ($m.Success) {
        $v = [int]$m.Groups[1].Value
        if ($v -ge [int]$c.our_version) { return $true }
        if ($v -ge [int]$c.other_tool_min_version) { return $true }
    }
    return $false
}

function Get-CloudMemFiles([string]$dataDir) {
    $d = Join-Path $dataDir 'memory'
    if (-not (Test-Path -LiteralPath $d)) { return @() }
    return @(Get-ChildItem -LiteralPath $d -Filter '*_memory.md' -File -ErrorAction SilentlyContinue)
}

function Set-ReadOnly([string]$file, [bool]$ro) {
    try { (Get-Item -LiteralPath $file -Force).IsReadOnly = $ro; return $true } catch { return $false }
}

function New-CloudMemText($cfg, [string]$uid, [string]$persona, [int]$version) {
    $now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    $t = [string]$cfg.block_template
    $t = $t.Replace('<iso8601>', $now).Replace('<version>', [string]$version)
    $t = $t.Replace('<persona_json_escaped>', (Get-JsonBody $persona))
    $t = $t.Replace('<persona>', $persona).Replace('<uid>', $uid)
    return (ConvertTo-Nl $t "`r`n")
}

function Get-UIDFromText([string]$file, [string]$text) {
    $m = [regex]::Match($text, '"uid"\s*:\s*"([^"]+)"')
    if ($m.Success) { return $m.Groups[1].Value }
    return ([System.IO.Path]::GetFileName($file) -replace '_memory\.md$', '')
}

function Invoke-CloudMem($ed, [string]$dataDir, [string]$action, [ref]$badRef) {
    $cfg = Get-CloudMemCfg
    $files = Get-CloudMemFiles $dataDir
    if ($action -eq 'status') {
        Write-Log ('-- 云记忆状态（' + $ed.displayName + '）：' + $dataDir)
        if ($files.Count -eq 0) { Write-Log '   (没有档案：先在客户端里发一句话生成云记忆)'; return }
        foreach ($f in $files) {
            $t = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
            $m = [regex]::Match($t, '(?s)## Memory Block\s*(.*?)\s*---')
            $len = 0
            if ($m.Success) { $len = ($m.Groups[1].Value).Trim().Length }
            $v = 0
            $mv = [regex]::Match($t, '(?m)^\s*>\s*Version:\s*(\d+)\s*$')
            if ($mv.Success) { $v = [int]$mv.Groups[1].Value }
            $ro = $false
            try { $ro = (Get-Item -LiteralPath $f.FullName -Force).IsReadOnly } catch {}
            Write-Log ('   ' + $f.Name + ' | 块 ' + $len + ' 字符 | Version ' + $v + ' | 只读 ' + $ro + ' | 本工具产物 ' + (Test-IsOurCloudMem $t))
        }
        return
    }
    if ($files.Count -eq 0) { Write-Log '云记忆：该账号还没有档案（先在客户端里发一句话生成），已跳过' 'Yellow'; return }

    if ($action -eq 'inject') {
        $persona = Get-PolicyBody (Resolve-PolicyLang '中')
        if (-not $persona) { $persona = Get-PolicyBody 'zh' }
        if (-not $persona) { Write-Log '云记忆：找不到政策文本' 'Red'; return }
        $max = [int]$cfg.limits.max_block_chars
        if ($persona.Length -gt $max) { Write-Log ('云记忆：政策正文 ' + $persona.Length + ' 字符超过上限 ' + $max) 'Red'; return }
        foreach ($f in $files) {
            $t = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
            $uid = Get-UIDFromText $f.FullName $t
            $bak = $f.FullName + [string]$cfg.backup_suffix
            if ((-not (Test-Path -LiteralPath $bak)) -and (-not (Test-IsOurCloudMem $t))) {
                try { Copy-Item -LiteralPath $f.FullName -Destination $bak -Force } catch {}
            }
            [void](Set-ReadOnly $f.FullName $false)
            Write-FileTracked $f.FullName (New-CloudMemText $cfg $uid $persona ([int]$cfg.our_version))
            if (-not (Set-ReadOnly $f.FullName $true)) { Write-Log ('   只读锁没设上（FAT32 或权限不足），不影响使用：' + $f.Name) 'Yellow' }
            Write-Log ('   已注入：' + $f.Name)
        }
        Write-Log ('云记忆注入完成：' + $files.Count + ' 个档案') 'Green'
        return
    }
    if ($action -eq 'revert') {
        foreach ($f in $files) {
            $bak = $f.FullName + [string]$cfg.backup_suffix
            $wasRo = $false
            try { $wasRo = (Get-Item -LiteralPath $f.FullName -Force).IsReadOnly } catch {}
            [void](Set-ReadOnly $f.FullName $false)
            $t = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
            $ours = Test-IsOurCloudMem $t
            $bakOk = $false
            if (Test-Path -LiteralPath $bak) {
                $bt = [System.IO.File]::ReadAllText($bak, [System.Text.Encoding]::UTF8)
                $bakOk = -not (Test-IsOurCloudMem $bt)
            }
            if ($bakOk) {
                Copy-Item -LiteralPath $bak -Destination $f.FullName -Force
                Remove-Item -LiteralPath $bak -Force -ErrorAction SilentlyContinue
                Write-Log ('   已还原：' + $f.Name)
            } elseif ($ours) {
                $uid = Get-UIDFromText $f.FullName $t
                Write-FileTracked $f.FullName (New-CloudMemText $cfg $uid '' 0)
                Write-Log ('   已清空注入内容：' + $f.Name + '（原内容已被覆盖，无法恢复——如实告知）') 'Yellow'
            } else {
                if ($wasRo) { [void](Set-ReadOnly $f.FullName $true) }
                Write-Log ('   跳过（不是本工具写的）：' + $f.Name)
            }
        }
        Write-Log '云记忆还原完成' 'Green'
    }
}

# ------------------------------------------------------------------ 锚点体检
function Invoke-DumpAnchor($dirs, $eds, [switch]$Json) {
    $result = New-Object System.Collections.ArrayList
    if (-not $Json) { Write-Log '=== 锚点体检（只读）===' }
    foreach ($k in @('cn', 'intl')) {
        $ed = Get-Ed $k
        if ($eds -and ($eds -notcontains $k)) { continue }
        if (-not $dirs.ContainsKey($k)) {
            if (-not $Json) { Write-Log ('-- ' + $ed.displayName + '：未安装/未找到') }
            continue
        }
        $install = $dirs[$k]
        if (-not $Json) { Write-Log ('-- ' + $ed.displayName + '：' + $install) }
        $edDataDir = Resolve-DataDir $ed
        # 与打补丁时用同一套靶点选择：旧版本插件缓存客户端根本不读，这里要是算进"目标"，
        # 打完补丁也会永远显示"待打"（那 18 个文件按设计就不打）——检测器会误报未生效。
        $edVer = ''
        try {
            $ll = Join-Path $edDataDir 'last-launch.json'
            if (Test-Path -LiteralPath $ll) { $edVer = [string](Read-Json $ll).version }
        } catch {}
        $edRow = [ordered]@{
            edition = $k
            displayName = $ed.displayName
            installDir = $install
            dataDir = $edDataDir
            clientVersion = $edVer
            patches = (New-Object System.Collections.ArrayList)
        }
        foreach ($p in (Get-AllPatches)) {
            if ($p.platforms -and ($p.platforms -notcontains $Platform)) { continue }
            $tgts = New-Object System.Collections.ArrayList
            $skipBefore = $script:StaleSkipped
            if ($p.locate) { foreach ($f in (Resolve-BundleTargets $install $p)) { [void]$tgts.Add($f) } }
            else { foreach ($f in (Resolve-ScopeTargets $install $edDataDir $ed $p.scope $edVer ([bool]$p.all_plugin_versions))) { [void]$tgts.Add($f) } }
            $skipNow = $script:StaleSkipped - $skipBefore
            $hit = 0; $done = 0
            $cpc = $null
            if ($p.op -eq 'content_policy_block') { $cpc = (Get-GroupWith 'content_policy').content_policy }
            $fgDoneMark = ''
            if ($p.op -eq 'replace_tag_block') {
                $rt = [System.IO.File]::ReadAllText((Join-Path (Join-Path $DataRoot 'patches') ([string]$p.replace_with_file)), [System.Text.Encoding]::UTF8)
                $fgDoneMark = ([regex]::Match($rt, '^\s*<[^>]+>')).Value
            }
            foreach ($f in $tgts) {
                try {
                    $t = [System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8)
                    # 整文件替换后原句已消失：先看 marker 判"已打"，再看 only_if_contains 判"该不该打"
                    $mkFile = [string]$p.marker
                    if ($p.op -eq 'replace_file' -and $mkFile -and $t.Contains($mkFile)) { $done++; continue }
                    if ($p.only_if_contains -and -not $t.Contains([string]$p.only_if_contains)) { continue }
                    switch ([string]$p.op) {
                        'insert_after' {
                            if ([regex]::IsMatch($t, [string]$p.anchor_re)) {
                                $hit++
                                $m0 = [regex]::Match($t, [string]$p.anchor_re)
                                $ins = $m0.Index + $m0.Length
                                $len = [Math]::Min(([string]$p.insert).Length, $t.Length - $ins)
                                if ($len -gt 0 -and $t.Substring($ins, $len) -eq [string]$p.insert) { $done++ }
                            }
                        }
                        'replace_literal' {
                            if ($t.Contains([string]$p.find)) { $hit++ }
                            elseif ($t.Contains([string]$p.replace)) { $done++ }
                        }
                        'replace_many' {
                            $anyHit = $false
                            foreach ($pr in $p.pairs) { if ($t.Contains([string]$pr[0])) { $anyHit = $true; break } }
                            if ($anyHit) { $hit++ }
                            else {
                                $allNew = $true
                                foreach ($pr in $p.pairs) { if (-not $t.Contains([string]$pr[1])) { $allNew = $false; break } }
                                if ($allNew) { $done++ }
                            }
                        }
                        'replace_file' {
                            $mk = [string]$p.marker
                            if (-not $mk) { $mk = [string]$p.verify_re }
                            if ($mk -and $t.Contains($mk)) { $done++ } else { $hit++ }
                        }
                        'replace_tag_block' {
                            $anyTag = $false
                            foreach ($tg in $p.tags) { if ($t.Contains('<' + $tg + '>')) { $anyTag = $true; break } }
                            if ($anyTag) { $hit++ }
                            if ($fgDoneMark -and $t.Contains($fgDoneMark)) { $done++ }
                        }
                        'content_policy_block' {
                            # 工程助手变体整段替换后没有 <content_policy>（这是设计），不算缺失
                            if ($t.Contains('The operator is an adult working on their own machines')) { $done++; break }
                            if (-not $t.Contains([string]$cpc.tag_open)) { break }
                            $hit++
                            $blk = Get-PolicyBlock $t ([string]$cpc.tag_open) ([string]$cpc.tag_close)
                            if ($blk) {
                                # JSON 里块是转义形态，必须解码后再比（与补丁路径同一套判据）
                                $cmp = $blk
                                if (-not ($blk.Contains("`n") -or $blk.Contains("`r"))) {
                                    try { $cmp = ('"' + $blk + '"') | ConvertFrom-Json } catch { $cmp = $blk }
                                }
                                $body = Get-PolicyBody (Resolve-PolicyLang $cmp)
                                if ($body -and (Test-SameBlock $cmp $body)) { $done++ }
                                else {
                                    # 另一档也算"已打"：用户可能刚用 11-强化档 切过档，
                                    # 检测器不传 -Level 时不能把强化档误判成"待打"
                                    $other = ''
                                    if ($Level -eq 'strong') { $other = 'standard' } else { $other = 'strong' }
                                    $body2 = Get-PolicyBody (Resolve-PolicyLang $cmp) $other
                                    if ($body2 -and (Test-SameBlock $cmp $body2)) { $done++ }
                                }
                            }
                        }
                        'replace_prompt_template' {
                            # 提示词存在才算这个补丁的靶点；已换成精简提示 = 已打
                            $pn = [string]$p.prompt_name
                            if ($t.Contains('"' + $pn + '"')) {
                                $mk = [string]$p.marker
                                if (-not $mk) { $mk = 'You are a local assistant running on the user' }
                                if ($t.Contains($mk)) { $done++ }
                                else { $hit++ }
                            }
                        }
                        default { $hit++ }
                    }
                } catch {}
            }
            $state = '无该块'
            if ($hit -gt 0) {
                if ($done -ge $hit) { $state = '已打' } else { $state = '待打' }
            } elseif ($done -gt 0) { $state = '已打' }
            if ($Json) {
                [void]$edRow.patches.Add([ordered]@{
                    id = [string]$p.id
                    targets = $tgts.Count
                    skipped = $skipNow
                    hit = $hit
                    done = $done
                    state = $state
                    defaultOn = ($p.enabled_by_default -ne $false)
                    optional = ($p.optional -eq $true)
                    platforms = @($p.platforms)
                    cliFlag = [string]$p.cli_flag
                })
            } else {
                $skipTxt = ''
                if ($skipNow -gt 0) { $skipTxt = ' | 跳过旧缓存 ' + $skipNow }
                $optTxt = ''
                if ($p.optional -eq $true) { $optTxt = '（可选：老客户端没有这段不算缺失）' }
                Write-Log ('   ' + ([string]$p.id).PadRight(20) + ' 目标 ' + $tgts.Count + $skipTxt + ' | 锚点命中 ' + $hit + ' | 已打 ' + $done + ' | ' + $state + $optTxt)
            }
        }
        if ($Json) { [void]$result.Add($edRow) }
    }
    if ($Json) {
        $payload = [ordered]@{
            schema = 'workbuddy-poja-state/1'
            toolVersion = [string](Get-EditionsCfg).toolVersion
            platform = $Platform
            generatedAt = (Get-Date).ToString('s')
            editions = @($result)
        }
        # 注意：PowerShell 变量名大小写不敏感，这里不能叫 $json —— 会和 [switch]$Json 撞成同一个变量，
        # 结果是把字符串强转成 SwitchParameter 直接报错（输出的 True 就是这么来的）。
        $jsonText = ($payload | ConvertTo-Json -Depth 8 -Compress)
        if ($JsonOut) {
            # 写文件：中文路径经 stdout 传回会因控制台代码页变乱码
            try {
                $dir = [System.IO.Path]::GetDirectoryName($JsonOut)
                if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                [System.IO.File]::WriteAllText($JsonOut, $jsonText, (New-Object System.Text.UTF8Encoding($false)))
                Write-Host ('<<<WBPOJA_JSONFILE>>>' + $JsonOut + '<<<WBPOJA_END>>>')
            } catch {
                Write-Host ('<<<WBPOJA_JSON>>>' + $jsonText + '<<<WBPOJA_END>>>')
            }
        } else {
            Write-Host ('<<<WBPOJA_JSON>>>' + $jsonText + '<<<WBPOJA_END>>>')
        }
    }
}

# ================================================================== 入口
if ($Help) {
    $helpVer = ''
    try { $helpVer = [string](Get-EditionsCfg).toolVersion } catch {}
    if ($helpVer) { Write-Host ('WorkBuddy 破甲工具包 v' + $helpVer + ' (Windows)') } else { Write-Host 'WorkBuddy 破甲工具包 (Windows)' }
    Write-Host @'

  install.ps1 -Mode Patch [-Edition auto|cn|intl] [-Level standard|strong]
  install.ps1 -Mode Restore

  可选: -WebFilter -NoFileGuard -AskMode -MinimalPrompt -EngineeringPrompt -ContractPrompt -NoEnv -NoDisk -KeepCache -NoJsGate
        -DumpAnchor [-AsJson]  锚点体检（-AsJson 输出机器可读 JSON）
        -PolicyLang auto|en|zh -InstallDir/-DataDir/-StateDir
        -DryRun -NoRestart -NoDialog -Silent
  独立: -Guard / -GuardRemove / -GuardStatus
        -CloudMem / -CloudMemRevert / -CloudMemStatus
        -DumpAnchor
        -DeepClean（还原的补救手段：没有任何干净备份时，按内容摘除可逆补丁）
        -PurgeStaleCaches（删除旧版本插件缓存里残留本工具改动的缓存目录）
'@
    exit 0
}

try { $null = Get-EditionsCfg } catch { Write-Host ('数据文件缺失：' + $_.Exception.Message) -ForegroundColor Red; exit 1 }

# 启动器现在是纯 ASCII 单行的 .bat（不设标题、不写中文横幅，避免行尾/编码被改就解析失败），
# 中文标题、版本号和"解除网络锁定"都在这里做。
$script:ToolVer = [string](Get-EditionsCfg).toolVersion
try { if ($Host.UI.RawUI) { $Host.UI.RawUI.WindowTitle = 'WorkBuddy 破甲工具包 v' + $script:ToolVer } } catch {}
if (-not $Silent) {
    try {
        Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -File -ErrorAction SilentlyContinue |
            Unblock-File -ErrorAction SilentlyContinue
    } catch {}
}

$specialMode = ($Guard -or $GuardRemove -or $GuardStatus -or $CloudMem -or $CloudMemRevert -or $CloudMemStatus -or $DumpAnchor -or $BlockUpdate -or $UnblockUpdate -or $AltFrontDoor -or $AltFrontDoorOff)

if ($Mode -eq 'Restore') { Invoke-RestoreAll; exit 0 }

# 探测客户端（特殊模式也要用）
$clientDirs = Find-ClientDirs
if ($InstallDir) {
    $forced = 'cn'
    if ($Edition -ne 'auto') { $forced = $Edition }
    elseif ($clientDirs.Count -eq 1) { $forced = @($clientDirs.Keys)[0] }
    $clientDirs[$forced] = $InstallDir
}

$edsRequested = @()
if ($Edition -ne 'auto') { $edsRequested = @($Edition) } else { $edsRequested = @('cn', 'intl') }

if ($Guard -or $GuardRemove) {
    if ($Guard) {
        $res = Install-GuardTasks $edsRequested
        foreach ($r in $res) {
            $state = '正常'
            if (-not $r.exists) { $state = '创建失败' }
            elseif (-not $r.wellformed) { $state = '结构异常' }
            elseif ($r.lastResult -ne $null -and (@(0, 267009, 267011) -notcontains $r.lastResult)) { $state = ('上次运行异常(' + $r.lastResult + ')') }
            Write-Log ('   ' + ([string]$r.name).PadRight(30) + $state + $(if ($r.why) { ' — ' + $r.why } else { '' }))
        }
        $bad = @($res | Where-Object { (-not $_.exists) -or (-not $_.wellformed) })
        if ($bad.Count -gt 0) {
            Write-Log '守护任务有异常，请用管理员身份重新运行。' 'Red'
            Show-Message 'WorkBuddy 守护任务' ('有 ' + $bad.Count + " 个任务没装好。`n`n请右键以管理员身份运行 18-安装守护任务.bat 再试一次。")
            exit 1
        }
        Show-Message 'WorkBuddy 守护任务' ('已安装 ' + $res.Count + ' 个守护任务（每 30 分钟 + 每次登录自动重打补丁）。')
    } else {
        $r = Remove-GuardTasks $edsRequested
        Write-Log ('已删除 ' + $r.removed + ' 个守护任务')
        if ($r.failed.Count -gt 0) { Write-Log ('删除失败：' + ($r.failed -join ', ')) 'Red' }
        Show-Message 'WorkBuddy 守护任务' ('已删除 ' + $r.removed + ' 个守护任务。')
    }
    exit 0
}

if ($GuardStatus) {
    foreach ($k in $edsRequested) {
        $ed = Get-Ed $k
        foreach ($n in @(($ed.guardTaskTag + '_Hourly'), ($ed.guardTaskTag + '_Logon'))) {
            $ex = Test-TaskExists $n
            $wf = $false; $lr = $null
            if ($ex) { $wf = (Test-TaskWellFormed $n).ok; $lr = Get-TaskLastResult $n }
            Write-Log ('   ' + ([string]$n).PadRight(30) + '存在 ' + $ex + ' | 结构正常 ' + $wf + ' | 上次结果 ' + $(if ($lr -eq $null) { 'n/a' } else { $lr }))
        }
    }
    exit 0
}

if ($DumpAnchor) { Invoke-DumpAnchor $clientDirs $edsRequested -Json:$AsJson; exit 0 }

if ($BlockUpdate -or $UnblockUpdate) {
    $eds = if ($Edition -ne 'auto') { @($Edition) } else { @('cn', 'intl') }
    $ok = 0; $bad = 0
    foreach ($k in $eds) {
        $inst = ''
        if ($clientDirs.ContainsKey($k)) { $inst = [string]$clientDirs[$k] }
        $done = $false
        if ($BlockUpdate) { $done = Invoke-UpdateBlock $k $inst $UpdateCachePath }
        else { $done = Invoke-UpdateUnblock $k $inst $UpdateCachePath }
        if ($done) { $ok++ } else { $bad++ }
    }
    if (-not $Silent) {
        if ($BlockUpdate) {
            Show-Message 'WorkBuddy 自动更新' ("已阻断 $ok 个版本。`n`n更新检查、提示和自动安装都不会再发生；已下载的安装包已移到备份目录。`n`n恢复：双击 14-恢复自动更新.bat，或 2-一键还原。")
        } else {
            Show-Message 'WorkBuddy 自动更新' ("已恢复 $ok 个版本。`n`n下次启动客户端会重新检查更新。")
        }
    }
    if ($bad -gt 0) { exit 1 }
    exit 0
}

if ($AltFrontDoor -or $AltFrontDoorOff) {
    # 备用前置门是 CN 专用：copilot.tencent.com 是国内版的备用 origin；
    # 国际版走 www.workbuddy.ai，套这个值会跨区，直接跳过。
    if ($Edition -eq 'intl') {
        Write-Log '  备用前置入口只适用于国内版（CN），国际版不做改动。' 'Yellow'
        exit 1
    }
    $eds = @('cn')
    $ok = 0; $bad = 0
    foreach ($k in $eds) {
        if ($Edition -eq 'auto' -and -not $clientDirs.ContainsKey($k)) { continue }
        $done = $false
        try { $done = Set-AltFrontDoor $k ([bool]$AltFrontDoor) } catch { $done = $false }
        if ($done) {
            $ok++
            if (-not $NoRestart -and -not $DryRun) { Restart-ClientEdition $k }
        } else { $bad++ }
    }
    if (-not $Silent) {
        if ($AltFrontDoor) {
            Show-Message 'WorkBuddy 备用前置入口' ("已为 $ok 个版本启用。`n`n内置模型的请求将发往 copilot.tencent.com/v2（客户端自己的备用前置门）；`n鉴权、配置、上报、自动更新仍走默认门。`n`n验证：发一条消息后看日志里的 url= 是否为 copilot.tencent.com。`n关闭：双击 17-恢复默认入口.bat。")
        } else {
            Show-Message 'WorkBuddy 备用前置入口' ("已为 $ok 个版本关闭，内置模型回到默认门。")
        }
    }
    if ($bad -gt 0) { exit 1 }
    exit 0
}

if ($PurgeStaleCaches) {
    $edList = @()
    if ($Edition -ne 'auto') { $edList = @($Edition) } else { $edList = @($clientDirs.Keys) }
    $v = ''
    foreach ($k in $edList) {
        if (-not $clientDirs.ContainsKey($k)) { continue }
        $ddP = Resolve-DataDir (Get-Ed $k)
        try {
            $ll = Join-Path $ddP 'last-launch.json'
            if (Test-Path -LiteralPath $ll) { $v = [string](Read-Json $ll).version }
        } catch {}
        $script:LogFile = Join-Path (Get-PatchRoot (Get-Ed $k)) (Get-Ed $k).logFileName
    }
    $n = Invoke-PurgeStaleCaches $edList $v
    Write-Log ('已删除陈旧插件缓存目录：' + $n + ' 个')
    Show-Message 'WorkBuddy 陈旧缓存清理' ('已删除 ' + $n + " 个旧版本插件缓存目录（里面残留了本工具的改动）。`n`n这些是客户端旧版本的插件缓存，当前版本不会读取；需要时客户端会自己重建。")
    exit 0
}

if ($DeepClean) {
    $bad = Invoke-DeepClean $edsRequested
    Show-Message 'WorkBuddy DeepClean' ("已按内容摘除可逆补丁（命令闸门 / 残留子句 / 网页过滤 / Ask 模式）。`n`n复查未生效项：" + $bad + "`n`n政策块与文件保护块没有确定原文，只能靠备份还原——若仍残留，请把 4-锚点体检.bat 的输出发给提供者。")
    exit 0
}

if ($CloudMem -or $CloudMemRevert -or $CloudMemStatus) {
    $act = 'status'
    if ($CloudMem) { $act = 'inject' }
    if ($CloudMemRevert) { $act = 'revert' }
    $badRef = 0
    foreach ($k in $edsRequested) {
        if (-not $clientDirs.ContainsKey($k)) { continue }
        $ed = Get-Ed $k
        $root = Get-PatchRoot $ed
        if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null }
        $script:LogFile = Join-Path $root $ed.logFileName
        Invoke-CloudMem $ed (Resolve-DataDir $ed) $act ([ref]$badRef)
    }
    $bad = Invoke-Recheck $edsRequested
    Show-Message 'WorkBuddy 云记忆' ('云记忆操作完成（' + $act + '）。' + $(if ($bad -gt 0) { "`n`n警告：有 $bad 处写入未生效。" } else { '' }))
    exit 0
}

# ---- 正常补丁流程
$targets = New-Object System.Collections.ArrayList
foreach ($k in $edsRequested) { if ($clientDirs.ContainsKey($k)) { [void]$targets.Add($k) } }

if ($targets.Count -eq 0) {
    $names = @()
    foreach ($k in @('cn', 'intl')) { $names += (Get-Ed $k).displayName }
    Show-Message 'WorkBuddy 一键破甲' ("没有找到任何客户端安装目录。`n`n请先安装并至少打开一次客户端，然后重新运行。`n`n已查找：" + ($names -join ' / '))
    exit 1
}

$report = @{}
foreach ($k in @('cn', 'intl')) {
    $report[$k] = @{
        displayName = (Get-Ed $k).displayName; ver = ''
        installDir = ''; dataDir = ''; snapshot = ''
        filesChanged = 0; envPrompts = 0; envName = ''; jsPatches = 0
        clauseHits = 0; guardFiles = 0; restarted = $false
        denied = 0; warnings = (New-Object System.Collections.ArrayList)
    }
}

# 两个版本装进同一个目录时，安装器会互相覆盖 resources/：
# 谁最后更新，另一个版本就被换成对方的界面与配置——表现为"更新后打不开 / 永远升不上去"。
# （2026-10-20 本机实测踩过：国内版目录被国际版 5.8.4 覆盖。）
if ($clientDirs.ContainsKey('cn') -and $clientDirs.ContainsKey('intl')) {
    $cnDir0 = [string]$clientDirs['cn']; $intlDir0 = [string]$clientDirs['intl']
    if ($cnDir0 -and $intlDir0 -and ($cnDir0.TrimEnd('\') -ieq $intlDir0.TrimEnd('\'))) {
        foreach ($k in @('cn', 'intl')) {
            [void]$report[$k].warnings.Add('两个版本装在同一目录：更新会互相覆盖，建议把国际版单独安装到别的文件夹')
        }
        Write-Log '  警告：国内版与国际版装在同一个目录，更新会互相覆盖（建议重装国际版到独立目录）' 'Yellow'
    }
}

# 提权重跑
if (-not $DryRun) {
    $needElevate = $false; $probeDir = ''
    foreach ($k in $targets) {
        $d = $clientDirs[$k]
        if (-not $d) { continue }
        $tpl = Join-Path $d (ConvertTo-PathSep (Get-EditionsCfg).paths.templates)
        if ((Test-Path -LiteralPath $tpl) -and (-not (Test-DirPatchable $tpl))) { $needElevate = $true; $probeDir = $d; break }
    }
    if ($needElevate -and -not $Elevated -and -not (Test-IsAdmin)) {
        Write-Log '安装目录需要管理员权限，正在请求提权重跑…'
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"'),
                     '-Mode', $Mode, '-Edition', $Edition, '-Level', $Level, '-PolicyLang', $PolicyLang,
                     '-InstallDir', ('"' + $probeDir + '"'), '-Elevated')
        if ($DataDir)     { $argList += @('-DataDir', ('"' + $DataDir + '"')) }
        if ($NoRestart)   { $argList += '-NoRestart' }
        if ($NoDialog)    { $argList += '-NoDialog' }
        if ($NoEnv)       { $argList += '-NoEnv' }
        if ($NoDisk)      { $argList += '-NoDisk' }
        if ($WebFilter)   { $argList += '-WebFilter' }
        if ($NoFileGuard) { $argList += '-NoFileGuard' }
        if ($AskMode)     { $argList += '-AskMode' }
        if ($MinimalPrompt) { $argList += '-MinimalPrompt' }
        if ($EngineeringPrompt) { $argList += '-EngineeringPrompt' }
        if ($ContractPrompt) { $argList += '-ContractPrompt' }
        if ($KeepCache)   { $argList += '-KeepCache' }
        if ($NoJsGate)    { $argList += '-NoJsGate' }
        try { Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs | Out-Null; exit 0 }
        catch { Write-Log ('提权被取消，继续以当前权限尝试：' + $_.Exception.Message) }
    }
}

foreach ($k in $targets) {
    $ed = Get-Ed $k
    $rep = $report[$k]
    $install = $clientDirs[$k]
    if (-not $install) { continue }
    $rep.installDir = $install

    $patchRoot = Get-PatchRoot $ed
    if (-not (Test-Path -LiteralPath $patchRoot)) { New-Item -ItemType Directory -Path $patchRoot -Force | Out-Null }
    $script:LogFile = Join-Path $patchRoot $ed.logFileName
    Write-Log ('================ ' + $ed.displayName + ' 开始（工具 v' + $script:ToolVer + '） ================')

    # 注意：PowerShell 变量名大小写不敏感，脚本参数 $DataDir 和这个局部变量是同一个名字。
    # 用 $dd 避免串味：否则处理完第一个版本后，第二个版本会被函数的 if ($DataDir) 直接返回上一版目录。
    $dd = Resolve-DataDir $ed
    $rep.dataDir = $dd
    try {
        $ll = Join-Path $dd 'last-launch.json'
        if (Test-Path -LiteralPath $ll) { $rep.ver = [string](Read-Json $ll).version }
    } catch {}
    Write-Log ('  安装目录：' + $install)
    Write-Log ('  数据目录：' + $dd)
    Write-Log ('  客户端版本：' + $(if ($rep.ver) { $rep.ver } else { '未知' }))

    # 只读/ACL 自动修复：客户机上常见“目录可写、现有 .tpl 只读”，Test-Writable 会漏判。
    # 这里在真正打补丁前再测一次；不可写就自动清只读，必要时用 takeown/icacls 授权。
    if (-not $DryRun) {
        $tplDir = Join-Path $install (ConvertTo-PathSep (Get-EditionsCfg).paths.templates)
        if ((Test-Path -LiteralPath $tplDir) -and (-not (Test-DirPatchable $tplDir))) {
            Write-Log '检测到模板文件只读或当前用户无写权限，正在自动修复（清只读 + 授权）…' 'Yellow'
            if (Repair-InstallPermissions $install) { Write-Log '  写入权限已修复。' 'Green' }
            else { Write-Log '  自动修复后仍不可写：可能是只读磁盘/网络盘/EDR 拦截；请把客户端搬到可写目录后重试。' 'Red' }
        }
    }

    if (-not $DryRun) {
        $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
        $script:SnapDir = Join-Path (Join-Path $patchRoot 'state') ($stamp + [string]$ed.snapshotSuffix)
        try {
            if (-not (Test-Path -LiteralPath $script:SnapDir)) { New-Item -ItemType Directory -Path $script:SnapDir -Force | Out-Null }
            Set-Content -LiteralPath (Join-Path $script:SnapDir 'manifest.tsv') -Value '' -Encoding UTF8
            $rep.snapshot = $script:SnapDir
            # 记下安装根：还原时用它把补丁靶点整体扫一遍（含"我们从未写过"的文件）
            Save-ManifestLine ('INSTALLROOT' + "`t" + $install)
        } catch { Write-Log ('无法创建快照目录：' + $_.Exception.Message) }
    } else { $script:SnapDir = $null }

    # 通道 A
    if (-not $NoEnv) {
        Write-Log '--- 通道 A：环境变量覆盖 ---'
        $pc = 0
        $r = Build-EnvConfig $ed $install $dd ([ref]$pc)
        if ($r.ok) {
            if ($DryRun) {
                Write-Log ('  [预览] 将写入 ' + $r.pathEnv + ' -> <' + $ed.envFileName + '>（' + $pc + ' 条提示词）')
                $rep.envPrompts = $pc
            } else {
                $envFile = Join-Path $patchRoot $ed.envFileName
                Write-FileTracked $envFile $r.json
                Set-EnvOwned $ed $r.pathEnv $envFile
                $rep.envName = $r.pathEnv
                $rep.envPrompts = $pc
                Write-Log ('  覆盖配置：' + $envFile + '（' + $pc + ' 条提示词）')
            }
        } else {
            Write-Log ('  跳过：' + $r.reason)
            if ($r.reason -eq 'no-source') { [void]$rep.warnings.Add('既没有运行时缓存也没有 cli/product.json，通道 A 跳过（磁盘通道仍生效）') }
            if ($r.reason -eq 'nothing-to-patch') { Write-Log ('  （数据源：' + $r.srcKind + '，其中没有需要改写的提示词）') }
        }
    } else { Write-Log '--- 通道 A：已按 -NoEnv 跳过 ---' }

    # 通道 B
    if (-not $NoDisk) {
        Write-Log '--- 通道 B：磁盘补丁 ---'
        # 按版本分别统计"跳过多少旧缓存"：这个计数器是脚本级的，不重置会把上一版的数量带过来
        $script:StaleSkipped = 0
        foreach ($p in (Get-AllPatches)) {
            if ($p.platforms -and ($p.platforms -notcontains $Platform)) { continue }
            if ($NoJsGate -and ([string]$p.id) -like 'gate.*') { Write-Log ('  ' + ([string]$p.id).PadRight(20) + ' 已按 -NoJsGate 跳过'); continue }
            if ($p.enabled_by_default -eq $false) {
                $on = $false
                if ($p.cli_flag -eq '--web-filter')    { $on = [bool]$WebFilter }
                if ($p.cli_flag -eq '--no-file-guard')  { $on = [bool]$NoFileGuard }
                if ($p.cli_flag -eq '--ask-mode')       { $on = [bool]$AskMode }
                if ($p.cli_flag -eq '--minimal-prompt') { $on = [bool]$MinimalPrompt }
                if ($p.cli_flag -eq '--engineering-prompt')   { $on = [bool]$EngineeringPrompt }
                if ($p.cli_flag -eq '--contract-prompt') { $on = [bool]$ContractPrompt }
                if (-not $on) { continue }
            }
            [void]$script:EnabledPatchIds.Add([string]$p.id)
            $tgts = New-Object System.Collections.ArrayList
            if ($p.locate) { foreach ($f in (Resolve-BundleTargets $install $p)) { [void]$tgts.Add($f) } }
            else { foreach ($f in (Resolve-ScopeTargets $install $dd $ed $p.scope $rep.ver ([bool]$p.all_plugin_versions))) { [void]$tgts.Add($f) } }
            if ($tgts.Count -eq 0) { continue }
            $changedFiles = 0; $hits = 0
            foreach ($f in $tgts) {
                $info = ''
                try {
                    $c = Invoke-PatchOnFile $f $p $ed ([ref]$info)
                    if ($c -gt 0) { $changedFiles++; $hits += $c }
                    if ($info -eq 'json-invalid') {
                        Write-Log ('  [跳过] JSON 改后无法解析，已放弃写入：' + $f) 'Red'
                        [void]$rep.warnings.Add((Split-Path -Leaf $f) + ' 自检失败，已放弃写入（未被改坏）')
                    }
                    if ($info -eq 'not-utf8') {
                        Write-Log ('  [跳过] 文件不是合法 UTF-8，改写有风险，已放弃：' + $f) 'Yellow'
                        [void]$rep.warnings.Add((Split-Path -Leaf $f) + ' 非合法 UTF-8，已跳过')
                    }
                    if ($info -eq 'length-mismatch') {
                        Write-Log ('  [跳过] 改动长度与预期不符，已放弃：' + $f) 'Red'
                        [void]$rep.warnings.Add((Split-Path -Leaf $f) + ' 改动长度异常，已跳过')
                    }
                    if ($info -eq 'access-denied') {
                        Write-Log ('  [跳过] 目标文件只读或 ACL 拒绝写入，已跳过：' + $f) 'Red'
                        [void]$rep.warnings.Add((Split-Path -Leaf $f) + ' 只读/无写权限，已跳过（安装器已尝试自动修复）')
                        $rep.denied++
                    }
                    if ($info -eq 'file-locked') {
                        Write-Log ('  [跳过] 目标文件被客户端占用，请完全退出客户端后重跑：' + $f) 'Yellow'
                        [void]$rep.warnings.Add((Split-Path -Leaf $f) + ' 被占用，已跳过')
                    }
                } catch {
                    Write-Log ('  [失败] ' + $f + ' : ' + $_.Exception.Message) 'Red'
                    [void]$rep.warnings.Add('跳过 ' + (Split-Path -Leaf $f))
                    if ($_.Exception.Message -match '拒绝访问|Access|denied|Unauthorized') { $rep.denied++ }
                }
            }
            if ($changedFiles -gt 0) {
                Write-Log ('  ' + ([string]$p.id).PadRight(20) + ' 改动 ' + $changedFiles + ' 个文件，命中 ' + $hits + ' 处' + $(if ($DryRun) { '（预览）' } else { '' }))
            }
            if ($p.id -like 'gate*' -or $p.id -eq 'mac.sandbox') { $rep.jsPatches += $changedFiles }
            if ($p.id -eq 'clauses' -or $p.id -eq 'askmode') { $rep.clauseHits += $hits }
            if ($p.id -eq 'fileguard') { $rep.guardFiles += $changedFiles }
            $rep.filesChanged += $changedFiles
        }
    } else { Write-Log '--- 通道 B：已按 -NoDisk 跳过 ---' }

    if ($script:StaleSkipped -gt 0) {
        Write-Log ('  跳过 ' + $script:StaleSkipped + ' 个插件缓存（旧版本，或客户端版本未知）')
        Write-Log '  —— 客户端不读它们；打了也只会留下无法还原的残留。如需清理：还原后按提示跑 -PurgeStaleCaches'
    }

    # ---------------- 运行时缓存
    $rcg = Get-GroupWith 'runtime_cache'
    if ($rcg -and -not $KeepCache -and -not $DryRun) {
        foreach ($pat in $rcg.runtime_cache.paths) {
            $c = ([string]$pat).Replace('<data>', $dd)
            # 支持通配（5.9.2 起国内版是 acc-product-config-v3-<pid>-<guid>.json）
            $files = New-Object System.Collections.ArrayList
            if ($c -match '[\*\?]') {
                $parent = [System.IO.Path]::GetDirectoryName($c)
                $leaf = [System.IO.Path]::GetFileName($c)
                if ($parent -and $leaf -and (Test-Path -LiteralPath $parent)) {
                    foreach ($f in @(Get-ChildItem -LiteralPath $parent -File -Filter $leaf -ErrorAction SilentlyContinue)) { [void]$files.Add($f.FullName) }
                }
            } elseif (Test-Path -LiteralPath $c) {
                [void]$files.Add($c)
            }
            foreach ($c in $files) {
                # 先备份再删：这是本工具唯一"只删不写"的动作，不备份就会留下无法还原的改动
                $backed = $false
                if ($script:SnapDir) {
                    try {
                        $flat = ($c -replace '[:\\/]', '_')
                        $dst = Join-Path $script:SnapDir ('deleted_' + $flat)
                        Copy-Item -LiteralPath $c -Destination $dst -Force -ErrorAction Stop
                        $sha = ''
                        try { $sha = (Get-FileHash -LiteralPath $c -Algorithm SHA256).Hash } catch {}
                        Save-ManifestLine ('FILEDEL' + "`t" + $c + "`t" + $dst + "`t" + $sha)
                        $backed = $true
                    } catch { Write-Log ('  缓存备份失败，跳过删除（不留下无法还原的改动）：' + (Split-Path -Leaf $c)) 'Yellow' }
                }
                if ($backed) {
                    try { Remove-Item -LiteralPath $c -Force -ErrorAction Stop; Write-Log ('  已清除运行时缓存（已备份，可还原）：' + (Split-Path -Leaf $c)) }
                    catch { Write-Log ('  缓存删除失败（不影响磁盘补丁）：' + (Split-Path -Leaf $c)) }
                }
            }
        }
    }

    # 重启
    if (-not $NoRestart -and -not $DryRun) {
        $exe = Resolve-ExePath $install $ed (Get-ProductMeta $install)
        if ($exe) {
            Write-Log ('--- 重启 ' + $ed.shortName + ' ---')
            $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ed.winProcNameFilter })
            foreach ($pr in $running) { try { $null = $pr.CloseMainWindow() } catch {} }
            if ($running.Count -gt 0) {
                $deadline = (Get-Date).AddSeconds(8)
                while ((Get-Date) -lt $deadline) {
                    Start-Sleep -Milliseconds 500
                    if (@(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ed.winProcNameFilter }).Count -eq 0) { break }
                }
                foreach ($pr in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ed.winProcNameFilter })) {
                    try { Stop-Process -Id $pr.Id -Force -ErrorAction Stop } catch {}
                }
                Start-Sleep -Seconds 1
            }
            try {
                $envFile = Join-Path $patchRoot $ed.envFileName
                # 必须用 ShellExecute 方式启动（cmd /c start），让新进程彻底脱离
                # 安装脚本的控制台。CreateNoWindow 只防"新开窗口"，不解除句柄继承：
                # 客户端一旦继承了这个 console，它活着窗口就一直不关。
                if ($rep.envName -and (Test-Path -LiteralPath $envFile)) {
                    [Environment]::SetEnvironmentVariable([string]$rep.envName, $envFile, 'Process')
                }
                $exeDir = [System.IO.Path]::GetDirectoryName($exe)
                $si = New-Object System.Diagnostics.ProcessStartInfo
                $si.FileName = $env:ComSpec
                $si.Arguments = '/c start "" /d "' + $exeDir + '" "' + $exe + '"'
                $si.UseShellExecute = $false
                $si.CreateNoWindow = $true
                [void][System.Diagnostics.Process]::Start($si)
                $rep.restarted = $true
                Write-Log '  已重新启动'
            } catch { Write-Log ('  启动失败：' + $_.Exception.Message) }
        }
    }
}

# 两个客户端读同一个环境变量名（productConfigPathEnv 都是 ACC_PRODUCT_CONFIG_PATH），
# 所以环境变量覆盖同一时刻只能服务一个版本；磁盘补丁对两版各自生效、互不影响。
$envTargets = New-Object System.Collections.ArrayList
foreach ($k in $targets) {
    if ($report[$k].envName) {
        [void]$envTargets.Add($report[$k].envName + ' -> ' + (Join-Path (Get-PatchRoot (Get-Ed $k)) (Get-Ed $k).envFileName))
    }
}
if ($envTargets.Count -gt 1) {
    $lastOne = [string]$envTargets[$envTargets.Count - 1]
    Write-Log ('  注意：多个版本都用环境变量覆盖，但它们读同一个变量名，当前生效的是最后一个：' + $lastOne) 'Yellow'
    foreach ($k in $targets) {
        [void]$report[$k].warnings.Add('环境变量覆盖与另一版本共用同一个变量名，当前生效的是最后写入的那份；磁盘补丁不受影响')
    }
}

$badTotal = 0
if (-not $DryRun) { $badTotal = Invoke-Recheck $targets }

$hasIntl = ($targets -contains 'intl')
$title = 'WorkBuddy 一键破甲' + $(if ($hasIntl) { ' / WorkBuddy AI' } else { '' })
$msg = $(if ($DryRun) { "预览模式（未做任何修改）`n`n" } else { "破解完成！`n`n" })
foreach ($k in $targets) {
    $rep = $report[$k]
    $msg += ('【' + $rep.displayName + '】' + $(if ($rep.ver) { ' ' + $rep.ver } else { '' }) + "`n")
    $msg += ('  安装目录：' + $rep.installDir + "`n")
    $msg += ($(if ($DryRun) { '  将修改文件：' } else { '  已修改文件：' }) + $rep.filesChanged + ' 个（JS 补丁 ' + $rep.jsPatches + ' / 子句 ' + $rep.clauseHits + " 处）`n")
    $msg += ('  提示词覆盖：' + $rep.envPrompts + " 条`n")
    if ($rep.restarted) { $msg += "  已自动重启`n" }
    foreach ($w in $rep.warnings) { $msg += ('  · ' + $w + "`n") }
    if ($rep.denied -gt 0) { $msg += ('  重要：有 ' + $rep.denied + " 个文件没有写入权限，请用管理员身份再跑一次。`n") }
    $msg += "`n"
}
if ($envTargets.Count -gt 0) { $msg += ('环境变量覆盖当前指向（两版共用同一变量名）：' + [string]$envTargets[$envTargets.Count - 1] + "`n`n") }
if ($badTotal -gt 0) { $msg += ('警告：有 ' + $badTotal + " 处写入后未生效，请查看日志。`n`n") }
$msg += '不想用了：双击 2-一键还原.bat'
Show-Message $title $msg
exit 0
