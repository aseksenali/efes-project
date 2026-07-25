param(
  [Parameter(Mandatory=$true)][string]$Pptx,
  [Parameter(Mandatory=$true)][string]$OutDir,
  [int]$Width = 1600
)
$ErrorActionPreference = "Stop"
$Height = [int]($Width * 9 / 16)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
Get-ChildItem "$OutDir\slide-*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

$pp = New-Object -ComObject PowerPoint.Application
try {
  $deck = $pp.Presentations.Open($Pptx, $true, $false, $false)  # ReadOnly, Untitled, WithWindow=false
  $n = $deck.Slides.Count
  for ($i = 1; $i -le $n; $i++) {
    $out = Join-Path $OutDir ("slide-{0:D2}.png" -f $i)
    $deck.Slides.Item($i).Export($out, "PNG", $Width, $Height)
    Write-Output $out
  }
  $deck.Close()
} finally {
  $pp.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}
