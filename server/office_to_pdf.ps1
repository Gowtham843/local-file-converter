param(
  [Parameter(Mandatory = $true)]
  [string]$InputPath,

  [Parameter(Mandatory = $true)]
  [string]$OutputPath
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
  param($Object)
  if ($null -ne $Object) {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
  }
}

$inputFullPath = [System.IO.Path]::GetFullPath($InputPath)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$extension = [System.IO.Path]::GetExtension($inputFullPath).ToLowerInvariant()

switch ($extension) {
  ".ppt" {
    $app = $null
    $presentation = $null
    try {
      $app = New-Object -ComObject PowerPoint.Application
      $presentation = $app.Presentations.Open($inputFullPath, $true, $true, $false)
      $presentation.SaveAs($outputFullPath, 32)
    } finally {
      if ($null -ne $presentation) { $presentation.Close() }
      if ($null -ne $app) { $app.Quit() }
      Release-ComObject $presentation
      Release-ComObject $app
    }
  }
  ".pptx" {
    $app = $null
    $presentation = $null
    try {
      $app = New-Object -ComObject PowerPoint.Application
      $presentation = $app.Presentations.Open($inputFullPath, $true, $true, $false)
      $presentation.SaveAs($outputFullPath, 32)
    } finally {
      if ($null -ne $presentation) { $presentation.Close() }
      if ($null -ne $app) { $app.Quit() }
      Release-ComObject $presentation
      Release-ComObject $app
    }
  }
  ".doc" {
    $app = $null
    $document = $null
    try {
      $app = New-Object -ComObject Word.Application
      $app.Visible = $false
      $document = $app.Documents.Open($inputFullPath, $false, $true)
      $document.ExportAsFixedFormat($outputFullPath, 17)
    } finally {
      if ($null -ne $document) { $document.Close($false) }
      if ($null -ne $app) { $app.Quit() }
      Release-ComObject $document
      Release-ComObject $app
    }
  }
  ".docx" {
    $app = $null
    $document = $null
    try {
      $app = New-Object -ComObject Word.Application
      $app.Visible = $false
      $document = $app.Documents.Open($inputFullPath, $false, $true)
      $document.ExportAsFixedFormat($outputFullPath, 17)
    } finally {
      if ($null -ne $document) { $document.Close($false) }
      if ($null -ne $app) { $app.Quit() }
      Release-ComObject $document
      Release-ComObject $app
    }
  }
  ".xls" {
    $app = $null
    $workbook = $null
    try {
      $app = New-Object -ComObject Excel.Application
      $app.Visible = $false
      $workbook = $app.Workbooks.Open($inputFullPath, 3, $true)
      $workbook.ExportAsFixedFormat(0, $outputFullPath)
    } finally {
      if ($null -ne $workbook) { $workbook.Close($false) }
      if ($null -ne $app) { $app.Quit() }
      Release-ComObject $workbook
      Release-ComObject $app
    }
  }
  ".xlsx" {
    $app = $null
    $workbook = $null
    try {
      $app = New-Object -ComObject Excel.Application
      $app.Visible = $false
      $workbook = $app.Workbooks.Open($inputFullPath, 3, $true)
      $workbook.ExportAsFixedFormat(0, $outputFullPath)
    } finally {
      if ($null -ne $workbook) { $workbook.Close($false) }
      if ($null -ne $app) { $app.Quit() }
      Release-ComObject $workbook
      Release-ComObject $app
    }
  }
  default {
    throw "Unsupported Office file type: $extension"
  }
}
