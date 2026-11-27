# 打一个干净的发布包（不含缓存、日志、测试文件）
# 用法：pwsh -File .\release.ps1
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$ver = (Get-Content (Join-Path $root "VERSION") -Raw).Trim()
if (-not $ver) { throw "VERSION 文件为空" }

$stage = Join-Path $env:TEMP "meegle-bulk-bug-v$ver"
$dist = Join-Path $root "dist"
$zip = Join-Path $dist "meegle-bulk-bug-v$ver.zip"

if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null
New-Item -ItemType Directory -Path $dist -Force | Out-Null

$files = @(
    "meegle_bulk.py",
    "start_gui.bat",
    "stop_gui.bat",
    "README.md",
    "CHANGELOG.md",
    "VERSION",
    "config.example.json"
)
foreach ($f in $files) {
    Copy-Item (Join-Path $root $f) (Join-Path $stage $f) -Force
}

# 程序与文档
Copy-Item (Join-Path $root "webapp") (Join-Path $stage "webapp") -Recurse -Force
Copy-Item (Join-Path $root "docs") (Join-Path $stage "docs") -Recurse -Force

# 清掉本机缓存 / 登录态 / Python 缓存
Get-ChildItem $stage -Recurse -Force -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force
Get-ChildItem $stage -Recurse -Force -File -Include ".cache_users.json", "*.pyc" -ErrorAction SilentlyContinue |
    Remove-Item -Force

"stage files: $((Get-ChildItem $stage -Recurse -File).Count)"
if (Test-Path $zip) { Remove-Item $zip -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($stage, $zip, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Remove-Item $stage -Recurse -Force

$hash = (Get-FileHash $zip -Algorithm SHA256).Hash
$size = [math]::Round((Get-Item $zip).Length / 1MB, 2)
"包已生成：$zip"
"版本：V$ver"
"大小：$size MB"
"SHA256：$hash"
"$hash  meegle-bulk-bug-v$ver.zip" | Set-Content (Join-Path $dist "SHA256.txt") -Encoding utf8
