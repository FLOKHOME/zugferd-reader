param([string]$Pdf,[string]$Out)
$ErrorActionPreference='Stop'
[void][Reflection.Assembly]::LoadWithPartialName('System')
$b=[IO.File]::ReadAllBytes($Pdf)
$enc=[Text.Encoding]::GetEncoding(28591)
$s=$enc.GetString($b)
function Is-Invoice([string]$t){
  return (($t.IndexOf('<rsm:CrossIndustryInvoice') -ge 0) -or ($t.IndexOf('<CrossIndustryInvoice') -ge 0) -or ($t.IndexOf('<CrossIndustryDocument') -ge 0))
}
$idx=0
while($true){
  $sp=$s.IndexOf('stream',$idx)
  if($sp -lt 0){break}
  if(($sp -ge 3) -and ($s.Substring($sp-3,3) -eq 'end')){$idx=$sp+6;continue}
  $j=$sp+6
  if(([int]$s[$j] -eq 13) -and ([int]$s[$j+1] -eq 10)){$j=$j+2}
  elseif(([int]$s[$j] -eq 10) -or ([int]$s[$j] -eq 13)){$j=$j+1}
  $ep=$s.IndexOf('endstream',$j)
  if($ep -lt 0){break}
  $seg=New-Object byte[] ($ep-$j)
  [Array]::Copy($b,$j,$seg,0,($ep-$j))
  $pre=$s.Substring(0,$sp)
  $hits=[regex]::Matches($pre,'[0-9]+\s+0\s+obj')
  if($hits.Count -gt 0){ $dict=$s.Substring($hits[$hits.Count-1].Index, $sp-$hits[$hits.Count-1].Index) }
  else { $dict=$s.Substring([Math]::Max(0,$sp-500), [Math]::Min(500,$sp)) }
  $hasFlate=$dict.IndexOf('/FlateDecode') -ge 0
  $cands=New-Object System.Collections.ArrayList
  [void]$cands.Add($seg)
  if($hasFlate){
    $off=0
    if(($seg.Length -gt 2) -and ($seg[0] -eq 0x78)){$off=2}
    try{
      $ms=New-Object IO.MemoryStream
      $ms.Write($seg,$off,$seg.Length-$off); $ms.Position=0
      $df=New-Object IO.Compression.DeflateStream($ms,[IO.Compression.CompressionMode]::Decompress)
      $ou=New-Object IO.MemoryStream
      $buf=New-Object byte[] 65536
      while(($n=$df.Read($buf,0,$buf.Length)) -gt 0){$ou.Write($buf,0,$n)}
      [void]$cands.Add($ou.ToArray())
      $df.Close();$ms.Close();$ou.Close()
    }catch{}
  }
  foreach($d in $cands){
    if($d -eq $null){continue}
    if(Is-Invoice ($enc.GetString($d))){
      [IO.File]::WriteAllBytes($Out,$d); Write-Output 'OK'; exit 0
    }
  }
  $idx=$ep+9
}
exit 2
