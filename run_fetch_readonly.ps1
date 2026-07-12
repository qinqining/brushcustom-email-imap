param(
    [Parameter(Mandatory = $true)]
    [string]$Start,

    [string]$End = ""
)

$ErrorActionPreference = "Stop"

$Python = "C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$ProjectDir = "D:\brushcustom-email-imap"
$Script = Join-Path $ProjectDir "fetch_imap.py"

if ($End -eq "") {
    & $Python $Script --start $Start
} else {
    & $Python $Script --start $Start --end $End
}
