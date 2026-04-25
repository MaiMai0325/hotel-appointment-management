Add-Type -AssemblyName System.Windows.Forms

$workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "森下ホテルのExcelを選択"
$dialog.Filter = "Excel Macro Workbook (*.xlsm)|*.xlsm|Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$dialog.InitialDirectory = $workspace

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
  Write-Output "Cancelled"
  exit 0
}

$selectedPath = $dialog.FileName
$copiedWorkbook = Join-Path $workspace "morishita-room-plan.xlsm"
$syncScript = Join-Path $workspace "sync-morishita-data.ps1"
$sitePath = Join-Path $workspace "index.html"

Copy-Item -LiteralPath $selectedPath -Destination $copiedWorkbook -Force
powershell -ExecutionPolicy Bypass -File $syncScript -WorkbookPath $copiedWorkbook

if ($LASTEXITCODE -ne 0) {
  throw "Sync failed."
}

Start-Process $sitePath
Write-Output "Synced and opened site."
