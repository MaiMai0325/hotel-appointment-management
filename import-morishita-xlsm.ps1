param(
  [Parameter(Mandatory = $true)]
  [string]$WorkbookPath,

  [int]$Year = 2026,

  [int]$Month = 4,

  [string]$OutputPath = "inventory-import.tsv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ColumnName {
  param([int]$ColumnNumber)

  $name = ""
  $current = $ColumnNumber
  while ($current -gt 0) {
    $remainder = ($current - 1) % 26
    $name = ([char][int](65 + $remainder)) + $name
    $current = [math]::Floor(($current - 1) / 26)
  }
  return $name
}

function Get-CellText {
  param(
    [hashtable]$CellMap,
    [string[]]$SharedStrings,
    [int]$RowNumber,
    [int]$ColumnNumber
  )

  $cellRef = "{0}{1}" -f (Get-ColumnName -ColumnNumber $ColumnNumber), $RowNumber
  if (-not $CellMap.ContainsKey($cellRef)) {
    return ""
  }

  $cell = $CellMap[$cellRef]
  $valueNode = $cell.SelectSingleNode("./x:v", $script:Ns)
  if ($null -eq $valueNode) {
    return ""
  }

  $raw = $valueNode.InnerText
  $type = $cell.GetAttribute("t")

  if ($type -eq "s") {
    return $SharedStrings[[int]$raw]
  }

  return $raw
}

function Get-CellStyleIndex {
  param(
    [hashtable]$CellMap,
    [int]$RowNumber,
    [int]$ColumnNumber
  )

  $cellRef = "{0}{1}" -f (Get-ColumnName -ColumnNumber $ColumnNumber), $RowNumber
  if (-not $CellMap.ContainsKey($cellRef)) {
    return 0
  }

  $style = $CellMap[$cellRef].GetAttribute("s")
  if ([string]::IsNullOrWhiteSpace($style)) {
    return 0
  }

  return [int]$style
}

function Convert-ExcelSerialToDate {
  param([double]$Serial)

  return [datetime]::FromOADate($Serial)
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)

try {
  $entries = @{}
  foreach ($entry in $zip.Entries) {
    $entries[$entry.FullName] = $entry
  }

  [xml]$workbookXml = (New-Object System.IO.StreamReader($entries["xl/workbook.xml"].Open())).ReadToEnd()
  [xml]$relsXml = (New-Object System.IO.StreamReader($entries["xl/_rels/workbook.xml.rels"].Open())).ReadToEnd()
  [xml]$sharedXml = (New-Object System.IO.StreamReader($entries["xl/sharedStrings.xml"].Open())).ReadToEnd()
  [xml]$stylesXml = (New-Object System.IO.StreamReader($entries["xl/styles.xml"].Open())).ReadToEnd()

  $script:Ns = New-Object System.Xml.XmlNamespaceManager($workbookXml.NameTable)
  $script:Ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  $script:Ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
  $sheetName = "{0}" -f $Year + [char]0x5E74
  $sheetNode = $workbookXml.SelectNodes("//x:sheets/x:sheet", $script:Ns) | Where-Object { $_.GetAttribute("name") -eq $sheetName } | Select-Object -First 1
  if ($null -eq $sheetNode) {
    throw ("Sheet not found: {0}" -f $sheetName)
  }

  $relId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
  $relNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
  $relNs.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")
  $relNode = $relsXml.SelectNodes("//r:Relationship", $relNs) | Where-Object { $_.GetAttribute("Id") -eq $relId } | Select-Object -First 1
  $sheetPath = "xl/" + $relNode.Target

  [xml]$sheetXml = (New-Object System.IO.StreamReader($entries[$sheetPath].Open())).ReadToEnd()
  $sheetNs = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
  $sheetNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

  $sharedStrings = @()
  foreach ($si in $sharedXml.SelectNodes("//x:si", $script:Ns)) {
    $parts = $si.SelectNodes(".//x:t", $script:Ns)
    $sharedStrings += (($parts | ForEach-Object { $_.InnerText }) -join "")
  }

  $fillIds = @()
  foreach ($xf in $stylesXml.SelectNodes("//x:cellXfs/x:xf", $script:Ns)) {
    $fillIds += [int]$xf.fillId
  }

  $cellMap = @{}
  foreach ($cell in $sheetXml.SelectNodes("//x:sheetData/x:row/x:c", $sheetNs)) {
    $cellMap[$cell.r] = $cell
  }

  $monthBlockStart = (($Month - 1) * 10) + 1
  $firstDaySerial = Get-CellText -CellMap $cellMap -SharedStrings $sharedStrings -RowNumber $monthBlockStart -ColumnNumber 3
  if ([string]::IsNullOrWhiteSpace($firstDaySerial)) {
    throw ("Month header not found for {0}-{1}" -f $Year, $Month)
  }

  $baseDate = Convert-ExcelSerialToDate -Serial ([double]$firstDaySerial)
  $roomRows = 4..9 | ForEach-Object { $monthBlockStart + ($_ - 1) }
  $daysInMonth = [datetime]::DaysInMonth($Year, $Month)
  $outputLines = New-Object System.Collections.Generic.List[string]

  foreach ($rowNumber in $roomRows) {
    $roomNumber = Get-CellText -CellMap $cellMap -SharedStrings $sharedStrings -RowNumber $rowNumber -ColumnNumber 1
    if ([string]::IsNullOrWhiteSpace($roomNumber)) {
      continue
    }

    for ($day = 1; $day -le $daysInMonth; $day++) {
      $columnNumber = $day + 2
      $cellText = Get-CellText -CellMap $cellMap -SharedStrings $sharedStrings -RowNumber $rowNumber -ColumnNumber $columnNumber
      $styleIndex = Get-CellStyleIndex -CellMap $cellMap -RowNumber $rowNumber -ColumnNumber $columnNumber
      $fillId = $fillIds[$styleIndex]

      $normalizedText = ($cellText -replace "\s", "")
      $isCheckoutOnly = $normalizedText -eq "退房"
      $isOccupied = ($fillId -ne 0) -and (-not $isCheckoutOnly)

      $date = Get-Date -Year $Year -Month $Month -Day $day -Format "yyyy-MM-dd"
      $stock = if ($isOccupied) { 0 } else { 1 }
      $outputLines.Add(("{0}`t{1}`t{2}" -f $date, $roomNumber, $stock))
    }
  }

  Set-Content -LiteralPath $OutputPath -Value $outputLines -Encoding utf8
  Write-Output ("Created: {0}" -f (Resolve-Path -LiteralPath $OutputPath))
  Write-Output ("Month parsed from workbook header: {0:yyyy-MM}" -f $baseDate)
}
finally {
  $zip.Dispose()
}
