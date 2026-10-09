param([string]$Base,[string]$Lines,[string]$Out)
$ErrorActionPreference='Stop'
$enc=[Text.Encoding]::GetEncoding(28591)
$la=[IO.File]::ReadAllLines($Lines,[Text.Encoding]::UTF8)

$C_BAND='0.13 0.27 0.42'; $C_ACC='0.85 0.55 0.15'
$C_HEAD='0.13 0.27 0.42'; $C_GREY='0.44 0.44 0.44'
$C_LINE='0.80 0.84 0.88'; $C_ZEBRA='0.93 0.95 0.97'
$C_PANEL='0.95 0.96 0.98'; $C_TOTBG='0.98 0.94 0.82'
$C_TOTX='0.55 0.24 0.03'; $C_BLACK='0 0 0'; $C_WHITE='1 1 1'
$X0=42.0; $X1=553.0; $W=511.0
$COLW=@(26,66,152,52,68,68,30,49)
$COLR=@(0,0,0,1,1,1,1,1)

function Esc([string]$t){ return $t.Replace('\','\\').Replace('(','\(').Replace(')','\)') }
function TW([string]$t,[double]$fs){ return ($t.Length*$fs*0.52) }
function Fit([string]$t,[int]$n){ if($t.Length -gt $n){return $t.Substring(0,$n-1)+'.'}; return $t }

$pages=New-Object System.Collections.ArrayList
$sb=New-Object Text.StringBuilder
$script:y=0.0
function NewPage(){
  [void]$pages.Add($sb.ToString())
  $script:sb=New-Object Text.StringBuilder
  [void]$sb.Append("$C_BAND rg 0 796 595 46 re f`n")
  [void]$sb.Append("$C_ACC rg 0 792 595 4 re f`n")
  [void]$sb.Append("BT /F2 10 Tf $C_WHITE rg $X0 811 Td (ZUGFeRD / Factur-X - Sichtkopie - Fortsetzung) Tj ET`n")
  $script:y=772.0
}
function Txt($x,$y,$fs,$font,$col,$t){ [void]$sb.Append("BT /$font $fs Tf $col rg $x $y Td ($(Esc $t)) Tj ET`n") }
function Rtxt($xe,$y,$fs,$font,$col,$t){ $x=$xe-(TW $t $fs); Txt $x $y $fs $font $col $t }
function Rect($x,$y,$w,$h,$col){ [void]$sb.Append("$col rg $x $y $w $h re f`n") }

# Kopfband
[void]$sb.Append("$C_BAND rg 0 760 595 82 re f`n")
[void]$sb.Append("$C_ACC rg 0 756 595 4 re f`n")
$title='ZUGFeRD / Factur-X Sichtkopie'; $sub=''
$first=$la[0].Split([char]9)
if($first[0] -eq 'TITLE'){ $title=$first[1]; if($first.Length -gt 2){$sub=$first[2]} }
Txt $X0 812 17 'F2' $C_WHITE $title
if($sub -ne ''){ Txt $X0 793 9 'F1' '0.86 0.90 0.94' (Fit $sub 95) }
$y=736.0
$alt=$false
foreach($ln in $la){
  if([string]::IsNullOrEmpty($ln)){continue}
  $p=$ln.Split([char]9)
  switch($p[0]){
    'TITLE'{}
    'META'{
      if($y -lt 70){NewPage}
      Txt $X0 $y 9 'F1' $C_GREY (Fit $p[1] 30)
      Txt ($X0+140) $y 9 'F1' $C_BLACK (Fit $p[2] 78)
      $y-=13.5
    }
    'HEAD'{
      if($y -lt 80){NewPage}
      $y-=9
      Txt $X0 $y 11 'F2' $C_HEAD $p[1]
      Rect $X0 ($y-4.5) $W 1.2 $C_ACC
      $y-=18
    }
    'TH'{
      if($y -lt 70){NewPage}
      Rect $X0 ($y-4) $W 16 $C_BAND
      $cx=$X0
      for($i=1;$i -lt $p.Count;$i++){
        $col=$p[$i]; $cw=$COLW[$i-1]
        if($COLR[$i-1] -eq 1){ Rtxt ($cx+$cw-3) $y 8 'F2' $C_WHITE $col } else { Txt ($cx+3) $y 8 'F2' $C_WHITE $col }
        $cx+=$cw
      }
      $y-=19; $alt=$false
    }
    'TR'{
      if($y -lt 60){NewPage}
      if($alt){ Rect $X0 ($y-4) $W 15 $C_ZEBRA }
      $cx=$X0
      for($i=1;$i -lt $p.Count;$i++){
        $col=$p[$i]; $cw=$COLW[$i-1]
        if($i-1 -eq 2){$col=Fit $col 38} elseif($i-1 -ne 2){$col=Fit $col 18}
        if($COLR[$i-1] -eq 1){ Rtxt ($cx+$cw-3) $y 8 'F1' $C_BLACK $col } else { Txt ($cx+3) $y 8 'F1' $C_BLACK $col }
        $cx+=$cw
      }
      $y-=15; $alt=(-not $alt)
    }
    'TOT'{
      if($y -lt 70){NewPage}
      Rect 285 ($y-4) 268 15 $C_PANEL
      Txt 293 $y 9 'F1' $C_GREY (Fit $p[1] 40)
      Rtxt ($X1-6) $y 9 'F1' $C_BLACK $p[2]
      $y-=15
    }
    'TOTB'{
      if($y -lt 70){NewPage}
      $y-=4
      Rect 285 ($y-6) 268 23 $C_TOTBG
      [void]$sb.Append("$C_TOTX RG 1 w 285 $($y-6) 268 23 re S`n")
      Txt 293 ($y+2) 12 'F2' $C_TOTX (Fit $p[1] 34)
      Rtxt ($X1-6) ($y+2) 12 'F2' $C_TOTX $p[2]
      $y-=28
    }
    'NOTE'{
      if($y -lt 60){NewPage}
      $y-=5
      Txt $X0 $y 8.5 'F1' $C_GREY (Fit $p[1] 105)
      $y-=13
    }
  }
}
[void]$pages.Add($sb.ToString())

# Fusszeile mit Seitenzahl
for($i=0;$i -lt $pages.Count;$i++){
  $ft=New-Object Text.StringBuilder
  [void]$ft.Append("BT /F1 8 Tf 0.55 0.55 0.55 rg $X0 26 Td (Sichtkopie der maschinenlesbaren Rechnungsdaten  -  Seite $($i+1) / $($pages.Count)) Tj ET`n")
  [void]$ft.Append($pages[$i])
  $pages[$i]=$ft.ToString()
}

$F1='<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>'
$F2='<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>'
function ResD($n1,$n2){ return "<< /Font << /F1 $n1 0 R /F2 $n2 0 R >> >>" }

if([string]::IsNullOrEmpty($Base)){
  $obj=@{}; $obj[1]='<< /Type /Catalog /Pages 2 0 R >>'; $obj[3]=$F1; $obj[4]=$F2
  $n=5; $kids=@()
  for($k=0;$k -lt $pages.Count;$k++){
    $cn=$n; $pn=$n+1; $n+=2; $st=$pages[$k]
    $obj[$cn]="<< /Length $($enc.GetByteCount($st)) >>`nstream`n$st`nendstream"
    $obj[$pn]="<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources $(ResD 3 4) /Contents $cn 0 R >>"
    $kids+="$pn 0 R"
  }
  $obj[2]="<< /Type /Pages /Kids [$($kids -join ' ')] /Count $($pages.Count) >>"
  $maxN=$n
  $o=New-Object Text.StringBuilder; [void]$o.Append("%PDF-1.4`n")
  $off=@{}
  for($i=1;$i -lt $maxN;$i++){ if($obj.ContainsKey($i)){ $off[$i]=$o.Length; [void]$o.Append("$i 0 obj`n"); [void]$o.Append($obj[$i]); [void]$o.Append("`nendobj`n") } }
  $xo=$o.Length
  [void]$o.Append("xref`n0 $maxN`n0000000000 65535 f `n")
  for($i=1;$i -lt $maxN;$i++){ if($off.ContainsKey($i)){ [void]$o.Append(("{0:0000000000} 00000 n `n" -f $off[$i])) } else { [void]$o.Append("0000000000 65535 f `n") } }
  [void]$o.Append("trailer`n<< /Size $maxN /Root 1 0 R >>`nstartxref`n$xo`n%%EOF`n")
  [IO.File]::WriteAllBytes($Out,$enc.GetBytes($o.ToString())); Write-Output 'OK-standalone'; exit 0
}

$b=[IO.File]::ReadAllBytes($Base); $s=$enc.GetString($b)
function Get-Seg([int]$num){ $m=[regex]::Match($s,'(?<![0-9])'+$num+'\s+0\s+obj'); if(-not $m.Success){return $null}; $st=$m.Index+$m.Length; $en=$s.IndexOf('endobj',$st); if($en -lt 0){return $null}; return $s.Substring($st,$en-$st) }
function Get-Dict([string]$seg){ if([string]::IsNullOrEmpty($seg)){return $null}; $a=$seg.IndexOf('<<'); if($a -lt 0){return $null}; $d=0;$i=$a; while($i -lt $seg.Length){ $two=if($i+2 -le $seg.Length){$seg.Substring($i,2)}else{''}; if($two -eq '<<'){$d++;$i+=2;continue}; if($two -eq '>>'){$d--;$i+=2; if($d -eq 0){return $seg.Substring($a,$i-$a)}; continue}; $i++ }; return $null }
$mRoot=[regex]::Match($s,'/Root\s+(\d+)\s+0\s+R'); if(-not $mRoot.Success){exit 11}
$cat=Get-Dict (Get-Seg ([int]$mRoot.Groups[1].Value)); if([string]::IsNullOrEmpty($cat)){exit 12}
$mP=[regex]::Match($cat,'/Pages\s+(\d+)\s+0\s+R'); if(-not $mP.Success){exit 13}
$pg=Get-Dict (Get-Seg ([int]$mP.Groups[1].Value)); if([string]::IsNullOrEmpty($pg)){exit 14}
$mK=[regex]::Match($pg,'/Kids\s*\[(.*?)\]',[Text.RegularExpressions.RegexOptions]::Singleline); if(-not $mK.Success){exit 15}
$mC=[regex]::Match($pg,'/Count\s+(\d+)'); $oldC=if($mC.Success){[int]$mC.Groups[1].Value}else{0}
$max=0; foreach($mm in [regex]::Matches($s,'(?<![0-9])(\d+)\s+0\s+obj')){ $v=[int]$mm.Groups[1].Value; if($v -gt $max){$max=$v} }

$app=New-Object Text.StringBuilder; $off=@{}
$fn1=$max+1; $fn2=$max+2
$off[$fn1]=$s.Length+$app.Length; [void]$app.Append("$fn1 0 obj`n$F1`nendobj`n")
$off[$fn2]=$s.Length+$app.Length; [void]$app.Append("$fn2 0 obj`n$F2`nendobj`n")
$cur=$max+3; $newKids=@(); $pageParent=$max+3+2*$pages.Count
for($k=0;$k -lt $pages.Count;$k++){
  $st=$pages[$k]; $cn=$cur; $pn=$cur+1; $cur+=2
  $off[$cn]=$s.Length+$app.Length
  [void]$app.Append("$cn 0 obj`n<< /Length $($enc.GetByteCount($st)) >>`nstream`n$st`nendstream`nendobj`n")
  $off[$pn]=$s.Length+$app.Length
  [void]$app.Append("$pn 0 obj`n<< /Type /Page /Parent $pageParent 0 R /MediaBox [0 0 595 842] /Resources $(ResD $fn1 $fn2) /Contents $cn 0 R >>`nendobj`n")
  $newKids+="$pn 0 R"
}
$newPgN=$cur; $newCatN=$cur+1
$newPg=$pg.Replace($mK.Groups[1].Value, $mK.Groups[1].Value+' '+($newKids -join ' '))
if($mC.Success){ $newPg=$newPg.Replace($mC.Groups[0].Value, "/Count $($oldC+$pages.Count)") }
$off[$newPgN]=$s.Length+$app.Length; [void]$app.Append("$newPgN 0 obj`n$newPg`nendobj`n")
$off[$newCatN]=$s.Length+$app.Length; [void]$app.Append("$newCatN 0 obj`n$($cat.Replace($mP.Groups[0].Value,"/Pages $newPgN 0 R"))`nendobj`n")
$xo=$s.Length+$app.Length
[void]$app.Append("xref`n0 1`n0000000000 65535 f `n")
$nums=$off.Keys | Sort-Object; $i=0
while($i -lt $nums.Count){ $j=$i; while($j+1 -lt $nums.Count -and $nums[$j+1] -eq $nums[$j]+1){$j++}; [void]$app.Append("$($nums[$i]) $($nums[$j]-$nums[$i]+1)`n"); for($k=$i;$k -le $j;$k++){ [void]$app.Append(("{0:0000000000} 00000 n `n" -f $off[$nums[$k]])) }; $i=$j+1 }
$oldX=[int]([regex]::Match($s,'startxref\s+(\d+)').Groups[1].Value)
[void]$app.Append("trailer`n<< /Size $($newCatN+1) /Root $newCatN 0 R /Prev $oldX >>`nstartxref`n$xo`n%%EOF`n")
[IO.File]::WriteAllBytes($Out,$enc.GetBytes($s+$app.ToString())); Write-Output 'OK-append'; exit 0
