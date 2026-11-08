param([string]$src, [string]$outdir)
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $true, $false, $false)
if (Test-Path $outdir) { Remove-Item -Recurse -Force $outdir }
New-Item -ItemType Directory -Path $outdir | Out-Null
$i = 1
foreach ($s in $pres.Slides) { $s.Export((Join-Path $outdir ("s{0:D2}.png" -f $i)), "PNG", 1600, 900); $i++ }
$pres.Close()
$ppt.Quit()
