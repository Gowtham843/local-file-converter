$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$BundledPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$BundledBin = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\bin"
$PopplerBin = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin"

if (Test-Path $BundledBin) {
  $env:PATH = "$BundledBin;$env:PATH"
}

if (Test-Path $PopplerBin) {
  $env:PATH = "$PopplerBin;$env:PATH"
}

if (Test-Path $BundledPython) {
  & $BundledPython (Join-Path $Root "server\converter_server.py")
} else {
  & python (Join-Path $Root "server\converter_server.py")
}
