# =============================================================================
#  fix-bat-crlf.ps1 —— 修复“解压后双击 bat 报一堆『不是内部或外部命令』”
#  （WorkBuddy AI 国际版工具包）
# =============================================================================
#  症状：双击 1-一键破甲.bat 后，窗口里刷出一串
#           'xxx' 不是内部或外部命令，也不是可运行的程序或批处理文件。
#        脚本根本没跑起来。
#  原因：包里的 .bat 是 Unix 换行（只有 LF 没有 CR）。cmd.exe 解析批处理时会错行，
#        把半行中文/多字节字符当成命令去执行，于是每一行都报一次那个错。
#        （中文越多、报错越多；纯英文的 bat 有时侥幸能跑，所以容易被忽略。）
#  处理：① 把本目录及子目录里的 .bat/.cmd 全部补成 CRLF（只插入 CR，不动内容）；
#        ② 顺便体检编码：含有当前系统 ANSI 代码页（简体中文 = 936/GBK）表示不了的
#           字符时报出来，这类文件在 cmd.exe 下同样会乱码出错。
# =============================================================================
[CmdletBinding()]
param(
    [string]$Root = $PSScriptRoot,
    [switch]$Verify      # 只体检不修改（发版前用）
)

$ErrorActionPreference = 'Stop'
Write-Host ''
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host '    WorkBuddy AI 国际版工具包 · BAT 体检（换行 + 编码）' -ForegroundColor Cyan
Write-Host '  ============================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "  目录：$Root"
if ($Verify) { Write-Host '  模式：只体检，不修改' -ForegroundColor Yellow }
Write-Host ''

$ansi = [System.Text.Encoding]::GetEncoding(936)
$fixed = 0; $ok = 0; $encBad = 0

foreach ($f in (Get-ChildItem -LiteralPath $Root -Recurse -File -Include *.bat, *.cmd -ErrorAction SilentlyContinue)) {
    $b = [System.IO.File]::ReadAllBytes($f.FullName)

    # ---- 1. 换行体检 ----
    $bare = 0
    for ($i = 0; $i -lt $b.Length; $i++) {
        if ($b[$i] -eq 10 -and ($i -eq 0 -or $b[$i - 1] -ne 13)) { $bare++ }
    }

    # ---- 2. 编码体检：GBK 能否逐字节还原（cmd.exe 就是按系统 ANSI/936 解释这些字节的）----
    $text  = $ansi.GetString($b)
    $rt    = $ansi.GetBytes($text)
    $encOk = ($rt.Length -eq $b.Length)
    if ($encOk) {
        for ($i = 0; $i -lt $b.Length; $i++) {
            if ($rt[$i] -ne $b[$i]) { $encOk = $false; break }
        }
    }

    if ($bare -eq 0) {
        $ok++
        if ($encOk) {
            Write-Host ("  [OK]    {0}" -f $f.Name) -ForegroundColor Green
        } else {
            $encBad++
            Write-Host ("  [ENC!]  {0}  —— 含 GBK(936) 表示不了的字符，cmd 下可能乱码" -f $f.Name) -ForegroundColor Red
        }
        continue
    }

    if ($Verify) {
        Write-Host ("  [BAD]   {0}  (有 {1} 个裸 LF)" -f $f.Name, $bare) -ForegroundColor Red
        continue
    }

    $out = [System.Collections.Generic.List[byte]]::new($b.Length + 64)
    for ($i = 0; $i -lt $b.Length; $i++) {
        if ($b[$i] -eq 10 -and ($i -eq 0 -or $b[$i - 1] -ne 13)) { $out.Add([byte]13) }
        $out.Add($b[$i])
    }
    [System.IO.File]::WriteAllBytes($f.FullName, $out.ToArray())
    $fixed++
    Write-Host ("  [FIXED] {0}  (补了 {1} 个 CR)" -f $f.Name, $bare) -ForegroundColor Yellow
}

Write-Host ''
if ($Verify) {
    if ($fixed -eq 0 -and $encBad -eq 0) {
        Write-Host '  体检通过：所有 bat/cmd 都是 CRLF，且编码可在 GBK 下往返。' -ForegroundColor Green
    } else {
        Write-Host ("  体检不通过：{0} 个换行有问题，{1} 个编码有问题。" -f $fixed, $encBad) -ForegroundColor Red
        exit 1
    }
} elseif ($fixed -gt 0) {
    Write-Host "  修好了 $fixed 个文件。现在再双击 1-一键破甲.bat 就能正常跑了。" -ForegroundColor Green
} else {
    Write-Host "  没有发现问题：本目录下 $ok 个 bat/cmd 都是正常的 CRLF 换行。" -ForegroundColor Green
}
if ($encBad -gt 0) {
    Write-Host "  编码提示：上面标 [ENC!] 的文件请改成 GBK 编码再发版。" -ForegroundColor Red
}
Write-Host ''
Write-Host '  如果还是报同样的错：说明包被别的程序又改回了 LF，或者你双击的不是本目录里的 bat。' -ForegroundColor DarkGray
Write-Host ''
if ($Host.Name -eq 'ConsoleHost' -and -not $env:FIX_NO_PAUSE) {
    Write-Host '  按回车键关闭本窗口...' -ForegroundColor DarkGray
    [void](Read-Host)
}
