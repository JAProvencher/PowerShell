Param (
  $Target = 'C:\PowerShell\*.*',
  $MinFrags = [Int]'2',
  $ClearLog = 1
)

$Title = 'Edefrag'
$host.ui.RawUI.WindowTitle = $Title
$TempDir = "$env:TEMP"
$Temp1 = "$TempDir\$Title-1.tmp"
$Temp2 = "$TempDir\$Title-2.tmp"
$LogFile = "$PSScriptRoot\$Title.txt"
$PSDefaultParameterValues['Out-File:Encoding'] = 'default'

If (-Not(Test-Path "$PSScriptRoot\Contig.exe")) {
  If ((Get-Command 'Contig.exe' -ErrorAction SilentlyContinue) -eq $null) 
  { 
    Write-Host 'Contig.exe not found in script directory or search path'
    Write-Host 'Download it from https://docs.microsoft.com/en-us/sysinternals/downloads/contig'
    Exit
  }
}

If ($ClearLog) {
  If (Test-Path $LogFile) {
    Remove-Item $LogFile
  }
}

$Language = Get-Culture

If ($Language -eq 'en-US') {
  $Str01 = 'Allocated Size'
  $Str02 = 'Allocated Size'
  $Msg01 = 'Searching for fragmented files...'
  $Msg02 = "Analyze the files with minimum $MinFrags fragments..."
}

If ($Language -eq 'it-IT') {
  $Str01 = 'Dimensione allocata'
  $Str02 = 'Dimensioni allocate'
  $Msg01 = 'Ricerca di file frammentati...'
  $Msg02 = "Analizza i file con minimo $MinFrags frammenti..."
}

Write-Host $Msg01

Cmd.exe /c "Chcp 1252 & Contig.exe -nobanner -a -s $Target >$Temp1"

Write-Host $Msg02

Function GetFragSizes($FileName){
  If (Test-Path $FileName) {
    fsutil file layout $FileName >$Temp2
    ForEach ($Line in Get-Content $Temp2) {
      If (($Line -match $Str01) -Or ($Line -match $Str02)) {
        $Size = $Line.Split(':')[1]
        $Sizes = $Sizes + $Size + ' ' 
      }
    }
    Return $Sizes
  }
  Else {
    Return 'File cannot be found'
  }
}

[System.IO.File]::ReadLines($Temp1, [System.Text.Encoding]::Default) | ForEach-Object {
  $Line = $_
  If ($Line -match ' fragments') {
    $Line = ($Line -Replace ' fragments','~fragments')
    $Items = $Line.Split(' ')
    $FragItem = $Items[$Items.Count-1]
    $FragCount = [Int]$FragItem.Split('~')[0]
    If ($FragCount -ge $MinFrags) {
      $FileName = $Line.Substring(0,$Line.Length - $FragItem.Length - 7)
      Echo $FileName
      $FragSizes = GetFragSizes($FileName)
      $Result = "$FileName | $FragCount | $FragSizes"
      Write-Host $Result
      Write-Output $Result >>$LogFile
    }
  }
}