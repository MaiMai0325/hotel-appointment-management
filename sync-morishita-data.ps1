param(
  [string]$WorkbookPath = ".\morishita-room-plan.xlsm",
  [int]$Year = 2026,
  [string]$JsOutputPath = "preloaded-data.js",
  [string]$TsvOutputPath = "inventory-all.tsv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$occupiedFillIds = @(2, 5, 9, 10)

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

function Get-CellRef {
  param(
    [int]$RowNumber,
    [int]$ColumnNumber
  )

  return "{0}{1}" -f (Get-ColumnName -ColumnNumber $ColumnNumber), $RowNumber
}

function Get-CellText {
  param(
    [hashtable]$CellMap,
    [string[]]$SharedStrings,
    [System.Xml.XmlNamespaceManager]$Ns,
    [int]$RowNumber,
    [int]$ColumnNumber
  )

  $cellRef = Get-CellRef -RowNumber $RowNumber -ColumnNumber $ColumnNumber
  if (-not $CellMap.ContainsKey($cellRef)) {
    return ""
  }

  $cell = $CellMap[$cellRef]
  $valueNode = $cell.SelectSingleNode("./x:v", $Ns)
  if ($null -eq $valueNode) {
    return ""
  }

  $raw = $valueNode.InnerText
  if ($cell.GetAttribute("t") -eq "s") {
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

  $cellRef = Get-CellRef -RowNumber $RowNumber -ColumnNumber $ColumnNumber
  if (-not $CellMap.ContainsKey($cellRef)) {
    return 0
  }

  $style = $CellMap[$cellRef].GetAttribute("s")
  if ([string]::IsNullOrWhiteSpace($style)) {
    return 0
  }

  return [int]$style
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

  $workbookNs = New-Object System.Xml.XmlNamespaceManager($workbookXml.NameTable)
  $workbookNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  $workbookNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

  $sheetName = "{0}" -f $Year + [char]0x5E74
  $sheetNode = $workbookXml.SelectNodes("//x:sheets/x:sheet", $workbookNs) | Where-Object { $_.GetAttribute("name") -eq $sheetName } | Select-Object -First 1
  if ($null -eq $sheetNode) {
    throw ("Sheet not found: {0}" -f $sheetName)
  }

  $relNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
  $relNs.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")
  $relId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
  $relNode = $relsXml.SelectNodes("//r:Relationship", $relNs) | Where-Object { $_.GetAttribute("Id") -eq $relId } | Select-Object -First 1
  $sheetPath = "xl/" + $relNode.Target

  [xml]$sheetXml = (New-Object System.IO.StreamReader($entries[$sheetPath].Open())).ReadToEnd()
  $sheetNs = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
  $sheetNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

  $sharedStrings = @()
  foreach ($si in $sharedXml.SelectNodes("//x:si", $workbookNs)) {
    $parts = $si.SelectNodes(".//x:t", $workbookNs)
    $sharedStrings += (($parts | ForEach-Object { $_.InnerText }) -join "")
  }

  $fillIds = @()
  foreach ($xf in $stylesXml.SelectNodes("//x:cellXfs/x:xf", $workbookNs)) {
    $fillIds += [int]$xf.fillId
  }

  $cellMap = @{}
  foreach ($cell in $sheetXml.SelectNodes("//x:sheetData/x:row/x:c", $sheetNs)) {
    $cellMap[$cell.r] = $cell
  }

  $inventory = New-Object System.Collections.Generic.List[object]
  $tsvLines = New-Object System.Collections.Generic.List[string]
  $rawStates = @{}

  foreach ($month in 1..12) {
    $monthBlockStart = (($month - 1) * 10) + 1
    $headerValue = Get-CellText -CellMap $cellMap -SharedStrings $sharedStrings -Ns $sheetNs -RowNumber $monthBlockStart -ColumnNumber 3
    if ([string]::IsNullOrWhiteSpace($headerValue)) {
      continue
    }

    $daysInMonth = [datetime]::DaysInMonth($Year, $month)
    $roomRows = 4..9 | ForEach-Object { $monthBlockStart + ($_ - 1) }

    foreach ($rowNumber in $roomRows) {
      $roomNumber = Get-CellText -CellMap $cellMap -SharedStrings $sharedStrings -Ns $sheetNs -RowNumber $rowNumber -ColumnNumber 1
      if ([string]::IsNullOrWhiteSpace($roomNumber)) {
        continue
      }

      foreach ($day in 1..$daysInMonth) {
        $columnNumber = $day + 2
        $cellText = Get-CellText -CellMap $cellMap -SharedStrings $sharedStrings -Ns $sheetNs -RowNumber $rowNumber -ColumnNumber $columnNumber
        $styleIndex = Get-CellStyleIndex -CellMap $cellMap -RowNumber $rowNumber -ColumnNumber $columnNumber
        $fillId = $fillIds[$styleIndex]
        $date = Get-Date -Year $Year -Month $month -Day $day -Format "yyyy-MM-dd"
        $normalizedText = ($cellText -replace "\s", "")
        $hasCheckIn = $normalizedText -like "*入住*"
        $hasCheckOut = $normalizedText -like "*退房*"
        $rawStates["{0}|{1}" -f $roomNumber, $date] = [pscustomobject]@{
          date = $date
          type = [string]$roomNumber
          fillId = $fillId
          text = $normalizedText
          hasCheckIn = $hasCheckIn
          hasCheckOut = $hasCheckOut
          isPotentiallyOccupied = ($occupiedFillIds -contains $fillId) -or $hasCheckIn -or $hasCheckOut
        }
      }
    }
  }

  $rawStates.Keys | Sort-Object | ForEach-Object {
    $current = $rawStates[$_]
    $previousDate = (Get-Date $current.date).AddDays(-1).ToString("yyyy-MM-dd")
    $nextDate = (Get-Date $current.date).AddDays(1).ToString("yyyy-MM-dd")
    $previous = $rawStates["{0}|{1}" -f $current.type, $previousDate]
    $next = $rawStates["{0}|{1}" -f $current.type, $nextDate]

    $continuesFromPrevious = ($null -ne $previous) -and $previous.isPotentiallyOccupied
    $continuesToNext = ($null -ne $next) -and $next.isPotentiallyOccupied
    $isCheckoutOnly = $current.hasCheckOut -and (-not $current.hasCheckIn)
    $isSameDayTurnover = $current.hasCheckOut -and $current.hasCheckIn

    $isOccupied = $false
    if ($current.hasCheckIn -or $isSameDayTurnover) {
      $isOccupied = $true
    }
    elseif ($isCheckoutOnly) {
      $isOccupied = $false
    }
    elseif ($current.isPotentiallyOccupied) {
      $isOccupied = $continuesToNext -or (($null -ne $previous) -and $previous.hasCheckIn -and (-not $current.hasCheckOut))
      if (-not $isOccupied -and $continuesFromPrevious -and $continuesToNext) {
        $isOccupied = $true
      }
    }

    $stock = if ($isOccupied) { 0 } else { 1 }
    $inventory.Add([pscustomobject]@{
      date = $current.date
      type = $current.type
      stock = $stock
    })
    $tsvLines.Add(("{0}`t{1}`t{2}" -f $current.date, $current.type, $stock))
  }

  $json = $inventory | ConvertTo-Json -Depth 3 -Compress
  $js = "window.MORISHITA_PRELOADED_DATA = { inventory: $json };"
  Set-Content -LiteralPath $JsOutputPath -Value $js -Encoding utf8
  Set-Content -LiteralPath $TsvOutputPath -Value $tsvLines -Encoding utf8

  Write-Output ("Created: {0}" -f (Resolve-Path -LiteralPath $JsOutputPath))
  Write-Output ("Created: {0}" -f (Resolve-Path -LiteralPath $TsvOutputPath))
}
finally {
  $zip.Dispose()
}
