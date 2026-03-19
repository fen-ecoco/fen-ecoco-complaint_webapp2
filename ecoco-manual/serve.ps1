param(
  [int]$Port = 5173
)

$ErrorActionPreference = "Stop"

Write-Host "Starting local web server on port $Port..."
Write-Host "Open links:"
Write-Host " - User manual:   http://localhost:$Port/user.html"
Write-Host " - Staff manual:  http://localhost:$Port/staff.html"
Write-Host " - Internal full: http://localhost:$Port/index.html"
Write-Host ""

python -m http.server $Port --directory (Split-Path -Parent $MyInvocation.MyCommand.Path)

