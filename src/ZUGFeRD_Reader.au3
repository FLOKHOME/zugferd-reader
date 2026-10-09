; =====================================================================================================
; ZUGFeRD Reader v1.6  -  AutoIt-Tool zum Lesen & Validieren von ZUGFeRD/Factur-X Rechnungs-XML
; Unterstützt: ZUGFeRD 1.x, 2.x (Factur-X), XRechnung (UBL-CII)
; =====================================================================================================
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListViewConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <ButtonConstants.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiListView.au3>
#include <GuiEdit.au3>
#include <GuiRichEdit.au3>
#include <ColorConstants.au3>
#include <FontConstants.au3>
#include <TabConstants.au3>
#include <DateTimeConstants.au3>

; =====================================================================================================
; GLOBALE NAMENSRAUM-MAPPINGS
; =====================================================================================================
Global $g_sNS_CII = "xmlns:rsm='urn:un:uncefact:data:standard:CrossIndustryInvoice:100' "
Global $g_sNS_RAM = "xmlns:ram='urn:un:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100' "
Global $g_sNS_UDT = "xmlns:udt='urn:un:uncefact:data:standard:UnqualifiedDataType:100' "
Global $g_sNS_QDT = "xmlns:qdt='urn:un:uncefact:data:standard:QualifiedDataType:100' "
Global $g_sNS_ZUG1 = "xmlns:zf='urn:ferd:CrossIndustryDocument:invoice:1p0' "
Global $g_sNS_XRE = "xmlns:sac='urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2' " & _
                     "xmlns:cbc='urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2' "
Global $g_sSelectionNS = $g_sNS_CII & $g_sNS_RAM & $g_sNS_UDT & $g_sNS_QDT & $g_sNS_ZUG1 & $g_sNS_XRE

; =====================================================================================================
; GUI
; =====================================================================================================
Global $g_hGUI, $g_hTab, $g_hListView, $g_hEdit_Validation, $g_hEdit_XML
Global $g_sLastFile = ""
Global $g_hStatus, $g_hBtnBrowse, $g_hBtnParse, $g_hBtnExport, $g_hBtnSicht
Global $g_sOrigPdf = ""
Global $g_hLineListview, $g_hSumEdit

; --- PDF-Helfer (PowerShell): extrahiert die eingebettete ZUGFeRD/Factur-X XML ---
; Verwendet .NET (System.IO.Compression.DeflateStream), das auf Win7/Vista vorhanden ist.
Global $g_sPDFHelper = _
    "param([string]$Pdf,[string]$Out)" & @CRLF & _
    "$ErrorActionPreference='Stop'" & @CRLF & _
    "[void][Reflection.Assembly]::LoadWithPartialName('System')" & @CRLF & _
    "$b=[IO.File]::ReadAllBytes($Pdf)" & @CRLF & _
    "$enc=[Text.Encoding]::GetEncoding(28591)" & @CRLF & _
    "$s=$enc.GetString($b)" & @CRLF & _
    "function Is-Invoice([string]$t){" & @CRLF & _
    "  return (($t.IndexOf('<rsm:CrossIndustryInvoice') -ge 0) -or ($t.IndexOf('<CrossIndustryInvoice') -ge 0) -or ($t.IndexOf('<CrossIndustryDocument') -ge 0))" & @CRLF & _
    "}" & @CRLF & _
    "$idx=0" & @CRLF & _
    "while($true){" & @CRLF & _
    "  $sp=$s.IndexOf('stream',$idx)" & @CRLF & _
    "  if($sp -lt 0){break}" & @CRLF & _
    "  if(($sp -ge 3) -and ($s.Substring($sp-3,3) -eq 'end')){$idx=$sp+6;continue}" & @CRLF & _
    "  $j=$sp+6" & @CRLF & _
    "  if(([int]$s[$j] -eq 13) -and ([int]$s[$j+1] -eq 10)){$j=$j+2}" & @CRLF & _
    "  elseif(([int]$s[$j] -eq 10) -or ([int]$s[$j] -eq 13)){$j=$j+1}" & @CRLF & _
    "  $ep=$s.IndexOf('endstream',$j)" & @CRLF & _
    "  if($ep -lt 0){break}" & @CRLF & _
    "  $seg=New-Object byte[] ($ep-$j)" & @CRLF & _
    "  [Array]::Copy($b,$j,$seg,0,($ep-$j))" & @CRLF & _
    "" & @CRLF & _
    "  # Das EIGENE Objekt-Dict bestimmen (letzter ""N 0 obj"" vor dem stream)" & @CRLF & _
    "  $pre=$s.Substring(0,$sp)" & @CRLF & _
    "  $hits=[regex]::Matches($pre,'[0-9]+\s+0\s+obj')" & @CRLF & _
    "  if($hits.Count -gt 0){ $dict=$s.Substring($hits[$hits.Count-1].Index, $sp-$hits[$hits.Count-1].Index) }" & @CRLF & _
    "  else { $dict=$s.Substring([Math]::Max(0,$sp-500), [Math]::Min(500,$sp)) }" & @CRLF & _
    "  $hasFlate=$dict.IndexOf('/FlateDecode') -ge 0" & @CRLF & _
    "" & @CRLF & _
    "  $cands=New-Object System.Collections.ArrayList" & @CRLF & _
    "  [void]$cands.Add($seg)                      # Roh (unkomprimiert)" & @CRLF & _
    "  if($hasFlate){" & @CRLF & _
    "    $off=0" & @CRLF & _
    "    if(($seg.Length -gt 2) -and ($seg[0] -eq 0x78)){$off=2}" & @CRLF & _
    "    try{" & @CRLF & _
    "      $ms=New-Object IO.MemoryStream" & @CRLF & _
    "      $ms.Write($seg,$off,$seg.Length-$off); $ms.Position=0" & @CRLF & _
    "      $df=New-Object IO.Compression.DeflateStream($ms,[IO.Compression.CompressionMode]::Decompress)" & @CRLF & _
    "      $ou=New-Object IO.MemoryStream" & @CRLF & _
    "      $buf=New-Object byte[] 65536" & @CRLF & _
    "      while(($n=$df.Read($buf,0,$buf.Length)) -gt 0){$ou.Write($buf,0,$n)}" & @CRLF & _
    "      [void]$cands.Add($ou.ToArray())" & @CRLF & _
    "      $df.Close();$ms.Close();$ou.Close()" & @CRLF & _
    "    }catch{}" & @CRLF & _
    "  }" & @CRLF & _
    "  foreach($d in $cands){" & @CRLF & _
    "    if($d -eq $null){continue}" & @CRLF & _
    "    if(Is-Invoice ($enc.GetString($d))){" & @CRLF & _
    "      [IO.File]::WriteAllBytes($Out,$d); Write-Output 'OK'; exit 0" & @CRLF & _
    "    }" & @CRLF & _
    "  }" & @CRLF & _
    "  $idx=$ep+9" & @CRLF & _
    "}" & @CRLF & _
    "exit 2" & @CRLF & _
    ""

Global $g_sSichtHelper = _
    "param([string]$Base,[string]$Lines,[string]$Out)" & @CRLF & _
    "$ErrorActionPreference='Stop'" & @CRLF & _
    "$enc=[Text.Encoding]::GetEncoding(28591)" & @CRLF & _
    "$la=[IO.File]::ReadAllLines($Lines,[Text.Encoding]::UTF8)" & @CRLF & _
    "" & @CRLF & _
    "$C_BAND='0.13 0.27 0.42'; $C_ACC='0.85 0.55 0.15'" & @CRLF & _
    "$C_HEAD='0.13 0.27 0.42'; $C_GREY='0.44 0.44 0.44'" & @CRLF & _
    "$C_LINE='0.80 0.84 0.88'; $C_ZEBRA='0.93 0.95 0.97'" & @CRLF & _
    "$C_PANEL='0.95 0.96 0.98'; $C_TOTBG='0.98 0.94 0.82'" & @CRLF & _
    "$C_TOTX='0.55 0.24 0.03'; $C_BLACK='0 0 0'; $C_WHITE='1 1 1'" & @CRLF & _
    "$X0=42.0; $X1=553.0; $W=511.0" & @CRLF & _
    "$COLW=@(26,66,152,52,68,68,30,49)" & @CRLF & _
    "$COLR=@(0,0,0,1,1,1,1,1)" & @CRLF & _
    "" & @CRLF & _
    "function Esc([string]$t){ return $t.Replace('\','\\').Replace('(','\(').Replace(')','\)') }" & @CRLF & _
    "function TW([string]$t,[double]$fs){ return ($t.Length*$fs*0.52) }" & @CRLF & _
    "function Fit([string]$t,[int]$n){ if($t.Length -gt $n){return $t.Substring(0,$n-1)+'.'}; return $t }" & @CRLF & _
    "" & @CRLF & _
    "$pages=New-Object System.Collections.ArrayList" & @CRLF & _
    "$sb=New-Object Text.StringBuilder" & @CRLF & _
    "$script:y=0.0" & @CRLF & _
    "function NewPage(){" & @CRLF & _
    "  [void]$pages.Add($sb.ToString())" & @CRLF & _
    "  $script:sb=New-Object Text.StringBuilder" & @CRLF & _
    "  [void]$sb.Append(""$C_BAND rg 0 796 595 46 re f`n"")" & @CRLF & _
    "  [void]$sb.Append(""$C_ACC rg 0 792 595 4 re f`n"")" & @CRLF & _
    "  [void]$sb.Append(""BT /F2 10 Tf $C_WHITE rg $X0 811 Td (ZUGFeRD / Factur-X - Sichtkopie - Fortsetzung) Tj ET`n"")" & @CRLF & _
    "  $script:y=772.0" & @CRLF & _
    "}" & @CRLF & _
    "function Txt($x,$y,$fs,$font,$col,$t){ [void]$sb.Append(""BT /$font $fs Tf $col rg $x $y Td ($(Esc $t)) Tj ET`n"") }" & @CRLF & _
    "function Rtxt($xe,$y,$fs,$font,$col,$t){ $x=$xe-(TW $t $fs); Txt $x $y $fs $font $col $t }" & @CRLF & _
    "function Rect($x,$y,$w,$h,$col){ [void]$sb.Append(""$col rg $x $y $w $h re f`n"") }" & @CRLF & _
    "" & @CRLF & _
    "# Kopfband" & @CRLF & _
    "[void]$sb.Append(""$C_BAND rg 0 760 595 82 re f`n"")" & @CRLF & _
    "[void]$sb.Append(""$C_ACC rg 0 756 595 4 re f`n"")" & @CRLF & _
    "$title='ZUGFeRD / Factur-X Sichtkopie'; $sub=''" & @CRLF & _
    "$first=$la[0].Split([char]9)" & @CRLF & _
    "if($first[0] -eq 'TITLE'){ $title=$first[1]; if($first.Length -gt 2){$sub=$first[2]} }" & @CRLF & _
    "Txt $X0 812 17 'F2' $C_WHITE $title" & @CRLF & _
    "if($sub -ne ''){ Txt $X0 793 9 'F1' '0.86 0.90 0.94' (Fit $sub 95) }" & @CRLF & _
    "$y=736.0" & @CRLF & _
    "$alt=$false" & @CRLF & _
    "foreach($ln in $la){" & @CRLF & _
    "  if([string]::IsNullOrEmpty($ln)){continue}" & @CRLF & _
    "  $p=$ln.Split([char]9)" & @CRLF & _
    "  switch($p[0]){" & @CRLF & _
    "    'TITLE'{}" & @CRLF & _
    "    'META'{" & @CRLF & _
    "      if($y -lt 70){NewPage}" & @CRLF & _
    "      Txt $X0 $y 9 'F1' $C_GREY (Fit $p[1] 30)" & @CRLF & _
    "      Txt ($X0+140) $y 9 'F1' $C_BLACK (Fit $p[2] 78)" & @CRLF & _
    "      $y-=13.5" & @CRLF & _
    "    }" & @CRLF & _
    "    'HEAD'{" & @CRLF & _
    "      if($y -lt 80){NewPage}" & @CRLF & _
    "      $y-=9" & @CRLF & _
    "      Txt $X0 $y 11 'F2' $C_HEAD $p[1]" & @CRLF & _
    "      Rect $X0 ($y-4.5) $W 1.2 $C_ACC" & @CRLF & _
    "      $y-=18" & @CRLF & _
    "    }" & @CRLF & _
    "    'TH'{" & @CRLF & _
    "      if($y -lt 70){NewPage}" & @CRLF & _
    "      Rect $X0 ($y-4) $W 16 $C_BAND" & @CRLF & _
    "      $cx=$X0" & @CRLF & _
    "      for($i=1;$i -lt $p.Count;$i++){" & @CRLF & _
    "        $col=$p[$i]; $cw=$COLW[$i-1]" & @CRLF & _
    "        if($COLR[$i-1] -eq 1){ Rtxt ($cx+$cw-3) $y 8 'F2' $C_WHITE $col } else { Txt ($cx+3) $y 8 'F2' $C_WHITE $col }" & @CRLF & _
    "        $cx+=$cw" & @CRLF & _
    "      }" & @CRLF & _
    "      $y-=19; $alt=$false" & @CRLF & _
    "    }" & @CRLF & _
    "    'TR'{" & @CRLF & _
    "      if($y -lt 60){NewPage}" & @CRLF & _
    "      if($alt){ Rect $X0 ($y-4) $W 15 $C_ZEBRA }" & @CRLF & _
    "      $cx=$X0" & @CRLF & _
    "      for($i=1;$i -lt $p.Count;$i++){" & @CRLF & _
    "        $col=$p[$i]; $cw=$COLW[$i-1]" & @CRLF & _
    "        if($i-1 -eq 2){$col=Fit $col 38} elseif($i-1 -ne 2){$col=Fit $col 18}" & @CRLF & _
    "        if($COLR[$i-1] -eq 1){ Rtxt ($cx+$cw-3) $y 8 'F1' $C_BLACK $col } else { Txt ($cx+3) $y 8 'F1' $C_BLACK $col }" & @CRLF & _
    "        $cx+=$cw" & @CRLF & _
    "      }" & @CRLF & _
    "      $y-=15; $alt=(-not $alt)" & @CRLF & _
    "    }" & @CRLF & _
    "    'TOT'{" & @CRLF & _
    "      if($y -lt 70){NewPage}" & @CRLF & _
    "      Rect 285 ($y-4) 268 15 $C_PANEL" & @CRLF & _
    "      Txt 293 $y 9 'F1' $C_GREY (Fit $p[1] 40)" & @CRLF & _
    "      Rtxt ($X1-6) $y 9 'F1' $C_BLACK $p[2]" & @CRLF & _
    "      $y-=15" & @CRLF & _
    "    }" & @CRLF & _
    "    'TOTB'{" & @CRLF & _
    "      if($y -lt 70){NewPage}" & @CRLF & _
    "      $y-=4" & @CRLF & _
    "      Rect 285 ($y-6) 268 23 $C_TOTBG" & @CRLF & _
    "      [void]$sb.Append(""$C_TOTX RG 1 w 285 $($y-6) 268 23 re S`n"")" & @CRLF & _
    "      Txt 293 ($y+2) 12 'F2' $C_TOTX (Fit $p[1] 34)" & @CRLF & _
    "      Rtxt ($X1-6) ($y+2) 12 'F2' $C_TOTX $p[2]" & @CRLF & _
    "      $y-=28" & @CRLF & _
    "    }" & @CRLF & _
    "    'NOTE'{" & @CRLF & _
    "      if($y -lt 60){NewPage}" & @CRLF & _
    "      $y-=5" & @CRLF & _
    "      Txt $X0 $y 8.5 'F1' $C_GREY (Fit $p[1] 105)" & @CRLF & _
    "      $y-=13" & @CRLF & _
    "    }" & @CRLF & _
    "  }" & @CRLF & _
    "}" & @CRLF & _
    "[void]$pages.Add($sb.ToString())" & @CRLF & _
    "" & @CRLF & _
    "# Fusszeile mit Seitenzahl" & @CRLF & _
    "for($i=0;$i -lt $pages.Count;$i++){" & @CRLF & _
    "  $ft=New-Object Text.StringBuilder" & @CRLF & _
    "  [void]$ft.Append(""BT /F1 8 Tf 0.55 0.55 0.55 rg $X0 26 Td (Sichtkopie der maschinenlesbaren Rechnungsdaten  -  Seite $($i+1) / $($pages.Count)) Tj ET`n"")" & @CRLF & _
    "  [void]$ft.Append($pages[$i])" & @CRLF & _
    "  $pages[$i]=$ft.ToString()" & @CRLF & _
    "}" & @CRLF & _
    "" & @CRLF & _
    "$F1='<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>'" & @CRLF & _
    "$F2='<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>'" & @CRLF & _
    "function ResD($n1,$n2){ return ""<< /Font << /F1 $n1 0 R /F2 $n2 0 R >> >>"" }" & @CRLF & _
    "" & @CRLF & _
    "if([string]::IsNullOrEmpty($Base)){" & @CRLF & _
    "  $obj=@{}; $obj[1]='<< /Type /Catalog /Pages 2 0 R >>'; $obj[3]=$F1; $obj[4]=$F2" & @CRLF & _
    "  $n=5; $kids=@()" & @CRLF & _
    "  for($k=0;$k -lt $pages.Count;$k++){" & @CRLF & _
    "    $cn=$n; $pn=$n+1; $n+=2; $st=$pages[$k]" & @CRLF & _
    "    $obj[$cn]=""<< /Length $($enc.GetByteCount($st)) >>`nstream`n$st`nendstream""" & @CRLF & _
    "    $obj[$pn]=""<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources $(ResD 3 4) /Contents $cn 0 R >>""" & @CRLF & _
    "    $kids+=""$pn 0 R""" & @CRLF & _
    "  }" & @CRLF & _
    "  $obj[2]=""<< /Type /Pages /Kids [$($kids -join ' ')] /Count $($pages.Count) >>""" & @CRLF & _
    "  $maxN=$n" & @CRLF & _
    "  $o=New-Object Text.StringBuilder; [void]$o.Append(""%PDF-1.4`n"")" & @CRLF & _
    "  $off=@{}" & @CRLF & _
    "  for($i=1;$i -lt $maxN;$i++){ if($obj.ContainsKey($i)){ $off[$i]=$o.Length; [void]$o.Append(""$i 0 obj`n""); [void]$o.Append($obj[$i]); [void]$o.Append(""`nendobj`n"") } }" & @CRLF & _
    "  $xo=$o.Length" & @CRLF & _
    "  [void]$o.Append(""xref`n0 $maxN`n0000000000 65535 f `n"")" & @CRLF & _
    "  for($i=1;$i -lt $maxN;$i++){ if($off.ContainsKey($i)){ [void]$o.Append((""{0:0000000000} 00000 n `n"" -f $off[$i])) } else { [void]$o.Append(""0000000000 65535 f `n"") } }" & @CRLF & _
    "  [void]$o.Append(""trailer`n<< /Size $maxN /Root 1 0 R >>`nstartxref`n$xo`n%%EOF`n"")" & @CRLF & _
    "  [IO.File]::WriteAllBytes($Out,$enc.GetBytes($o.ToString())); Write-Output 'OK-standalone'; exit 0" & @CRLF & _
    "}" & @CRLF & _
    "" & @CRLF & _
    "$b=[IO.File]::ReadAllBytes($Base); $s=$enc.GetString($b)" & @CRLF & _
    "function Get-Seg([int]$num){ $m=[regex]::Match($s,'(?<![0-9])'+$num+'\s+0\s+obj'); if(-not $m.Success){return $null}; $st=$m.Index+$m.Length; $en=$s.IndexOf('endobj',$st); if($en -lt 0){return $null}; return $s.Substring($st,$en-$st) }" & @CRLF & _
    "function Get-Dict([string]$seg){ if([string]::IsNullOrEmpty($seg)){return $null}; $a=$seg.IndexOf('<<'); if($a -lt 0){return $null}; $d=0;$i=$a; while($i -lt $seg.Length){ $two=if($i+2 -le $seg.Length){$seg.Substring($i,2)}else{''}; if($two -eq '<<'){$d++;$i+=2;continue}; if($two -eq '>>'){$d--;$i+=2; if($d -eq 0){return $seg.Substring($a,$i-$a)}; continue}; $i++ }; return $null }" & @CRLF & _
    "$mRoot=[regex]::Match($s,'/Root\s+(\d+)\s+0\s+R'); if(-not $mRoot.Success){exit 11}" & @CRLF & _
    "$cat=Get-Dict (Get-Seg ([int]$mRoot.Groups[1].Value)); if([string]::IsNullOrEmpty($cat)){exit 12}" & @CRLF & _
    "$mP=[regex]::Match($cat,'/Pages\s+(\d+)\s+0\s+R'); if(-not $mP.Success){exit 13}" & @CRLF & _
    "$pg=Get-Dict (Get-Seg ([int]$mP.Groups[1].Value)); if([string]::IsNullOrEmpty($pg)){exit 14}" & @CRLF & _
    "$mK=[regex]::Match($pg,'/Kids\s*\[(.*?)\]',[Text.RegularExpressions.RegexOptions]::Singleline); if(-not $mK.Success){exit 15}" & @CRLF & _
    "$mC=[regex]::Match($pg,'/Count\s+(\d+)'); $oldC=if($mC.Success){[int]$mC.Groups[1].Value}else{0}" & @CRLF & _
    "$max=0; foreach($mm in [regex]::Matches($s,'(?<![0-9])(\d+)\s+0\s+obj')){ $v=[int]$mm.Groups[1].Value; if($v -gt $max){$max=$v} }" & @CRLF & _
    "" & @CRLF & _
    "$app=New-Object Text.StringBuilder; $off=@{}" & @CRLF & _
    "$fn1=$max+1; $fn2=$max+2" & @CRLF & _
    "$off[$fn1]=$s.Length+$app.Length; [void]$app.Append(""$fn1 0 obj`n$F1`nendobj`n"")" & @CRLF & _
    "$off[$fn2]=$s.Length+$app.Length; [void]$app.Append(""$fn2 0 obj`n$F2`nendobj`n"")" & @CRLF & _
    "$cur=$max+3; $newKids=@(); $pageParent=$max+3+2*$pages.Count" & @CRLF & _
    "for($k=0;$k -lt $pages.Count;$k++){" & @CRLF & _
    "  $st=$pages[$k]; $cn=$cur; $pn=$cur+1; $cur+=2" & @CRLF & _
    "  $off[$cn]=$s.Length+$app.Length" & @CRLF & _
    "  [void]$app.Append(""$cn 0 obj`n<< /Length $($enc.GetByteCount($st)) >>`nstream`n$st`nendstream`nendobj`n"")" & @CRLF & _
    "  $off[$pn]=$s.Length+$app.Length" & @CRLF & _
    "  [void]$app.Append(""$pn 0 obj`n<< /Type /Page /Parent $pageParent 0 R /MediaBox [0 0 595 842] /Resources $(ResD $fn1 $fn2) /Contents $cn 0 R >>`nendobj`n"")" & @CRLF & _
    "  $newKids+=""$pn 0 R""" & @CRLF & _
    "}" & @CRLF & _
    "$newPgN=$cur; $newCatN=$cur+1" & @CRLF & _
    "$newPg=$pg.Replace($mK.Groups[1].Value, $mK.Groups[1].Value+' '+($newKids -join ' '))" & @CRLF & _
    "if($mC.Success){ $newPg=$newPg.Replace($mC.Groups[0].Value, ""/Count $($oldC+$pages.Count)"") }" & @CRLF & _
    "$off[$newPgN]=$s.Length+$app.Length; [void]$app.Append(""$newPgN 0 obj`n$newPg`nendobj`n"")" & @CRLF & _
    "$off[$newCatN]=$s.Length+$app.Length; [void]$app.Append(""$newCatN 0 obj`n$($cat.Replace($mP.Groups[0].Value,""/Pages $newPgN 0 R""))`nendobj`n"")" & @CRLF & _
    "$xo=$s.Length+$app.Length" & @CRLF & _
    "[void]$app.Append(""xref`n0 1`n0000000000 65535 f `n"")" & @CRLF & _
    "$nums=$off.Keys | Sort-Object; $i=0" & @CRLF & _
    "while($i -lt $nums.Count){ $j=$i; while($j+1 -lt $nums.Count -and $nums[$j+1] -eq $nums[$j]+1){$j++}; [void]$app.Append(""$($nums[$i]) $($nums[$j]-$nums[$i]+1)`n""); for($k=$i;$k -le $j;$k++){ [void]$app.Append((""{0:0000000000} 00000 n `n"" -f $off[$nums[$k]])) }; $i=$j+1 }" & @CRLF & _
    "$oldX=[int]([regex]::Match($s,'startxref\s+(\d+)').Groups[1].Value)" & @CRLF & _
    "[void]$app.Append(""trailer`n<< /Size $($newCatN+1) /Root $newCatN 0 R /Prev $oldX >>`nstartxref`n$xo`n%%EOF`n"")" & @CRLF & _
    "[IO.File]::WriteAllBytes($Out,$enc.GetBytes($s+$app.ToString())); Write-Output 'OK-append'; exit 0" & @CRLF & _
    ""

_MainGUI()

Func _MainGUI()
    ; Fenster
    $g_hGUI = GUICreate("ZUGFeRD Reader v1.6", 920, 700, -1, -1, $WS_OVERLAPPEDWINDOW)
    GUISetFont(9.5, 400, 0, "Segoe UI")

    ; Menü
    Local $hMenu = GUICtrlCreateMenu("&Datei")
    Local $hItemOpen = GUICtrlCreateMenuItem("Öffnen...", $hMenu)
    Local $hItemSep = GUICtrlCreateMenuItem("", $hMenu)
    Local $hItemExit = GUICtrlCreateMenuItem("Beenden", $hMenu)
    Local $hHelp = GUICtrlCreateMenu("&Hilfe")
    Local $hItemAbout = GUICtrlCreateMenuItem("Info", $hHelp)

    ; Toolbar-Buttons
    $g_hBtnBrowse = GUICtrlCreateButton("📂 Öffnen", 10, 8, 105, 28)
    $g_hBtnParse  = GUICtrlCreateButton("▶ Parsen", 121, 8, 80, 28)
    $g_hBtnSicht  = GUICtrlCreateButton("🖨 Sichtkopie", 207, 8, 125, 28)
    $g_hBtnExport = GUICtrlCreateButton("💾 TXT", 338, 8, 70, 28)

    GUICtrlCreateLabel("Datei:", 414, 14, 30, 20)
    $g_hStatus = GUICtrlCreateLabel("Keine Datei geladen", 448, 14, 455, 20)

    ; Tabs
    $g_hTab = GUICtrlCreateTab(10, 44, 900, 645)

    ; --- TAB 1: Felder ---
    GUICtrlCreateTabItem("Ausgelesene Felder")
    $g_hListView = GUICtrlCreateListView("Feld|Wert", 20, 72, 870, 250, _
        BitOR($LVS_REPORT, $LVS_SINGLESEL, $LVS_NOSORTHEADER))
    _GUICtrlListView_SetColumnWidth($g_hListView, 0, 280)
    _GUICtrlListView_SetColumnWidth($g_hListView, 1, 560)
    GUICtrlCreateLabel("Positionen (Line Items):", 20, 330, 200, 20)
    $g_hLineListview = GUICtrlCreateListView("Pos.|Artikel-Nr.|Bezeichnung|Menge|Einheit|Einzelpreis|Netto|Steuer%|Steuer", _
        20, 350, 870, 200, BitOR($LVS_REPORT, $LVS_SINGLESEL, $LVS_NOSORTHEADER))
    _GUICtrlListView_SetColumnWidth($g_hLineListview, 0, 40)
    _GUICtrlListView_SetColumnWidth($g_hLineListview, 1, 100)
    _GUICtrlListView_SetColumnWidth($g_hLineListview, 2, 260)
    _GUICtrlListView_SetColumnWidth($g_hLineListview, 3, 60)
    _GUICtrlListView_SetColumnWidth($g_hLineListview, 4, 55)
    _GUICtrlListView_SetColumnWidth($g_hLineListview, 5, 90)
    _GUICtrlListView_SetColumnWidth($g_hLineListview, 6, 90)
    _GUICtrlListView_SetColumnWidth($g_hLineListview, 7, 55)
    _GUICtrlListView_SetColumnWidth($g_hLineListview, 8, 90)
    GUICtrlCreateLabel("Summen:", 20, 556, 80, 20)
    $g_hSumEdit = GUICtrlCreateEdit("", 20, 576, 870, 100, BitOR($ES_READONLY, $WS_VSCROLL))
    GUICtrlSetBkColor($g_hSumEdit, 0xF0F8FF)  ; AliceBlue

    ; --- TAB 2: Validierung ---
    GUICtrlCreateTabItem("Validierung / Prüfung")
    $g_hEdit_Validation = GUICtrlCreateEdit("", 20, 72, 870, 600, BitOR($ES_READONLY, $WS_VSCROLL))
    GUICtrlSetFont($g_hEdit_Validation, 9.5, 400, 0, "Consolas")

    ; --- TAB 3: Roh-XML ---
    GUICtrlCreateTabItem("Roh-XML (Auszug)")
    $g_hEdit_XML = GUICtrlCreateEdit("", 20, 72, 870, 600, BitOR($ES_READONLY, $WS_VSCROLL))
    GUICtrlSetFont($g_hEdit_XML, 9.5, 400, 0, "Consolas")

    GUICtrlCreateTabItem("")  ; Tab-Ende

    ; Dateiauswahl über Button/Menü ODER Drag & Drop (PDF und XML)
    GUICtrlSetState($g_hBtnParse, $GUI_DISABLE)
    GUICtrlSetState($g_hBtnExport, $GUI_DISABLE)
    GUISetState(@SW_SHOW)

    ; --- Drag & Drop aktivieren (nur DllCall auf shell32, keine WinAPIEx-UDFs nötig) ---
    DllCall("shell32.dll", "none", "DragAcceptFiles", "hwnd", $g_hGUI, "int", 1)
    GUIRegisterMsg($WM_DROPFILES, "WM_DROPFILES")

    While 1
        Local $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $GUI_EVENT_CLOSE, $hItemExit
                GUIDelete($g_hGUI)
                Exit
            Case $hItemOpen, $g_hBtnBrowse
                _OpenFile()
            Case $g_hBtnParse
                _ParseAndDisplay()
            Case $g_hBtnExport
                _ExportToText()
            Case $g_hBtnSicht
                _MakeSichtkopie()
            Case $hItemAbout
                MsgBox($MB_ICONINFORMATION, "Über ZUGFeRD Reader", _
                    "ZUGFeRD Reader v1.6" & @CRLF & _
                    "Liest Factur-X / ZUGFeRD 1.x/2.x / XRechnung" & @CRLF & _
                    "Direkt aus PDF (Drag & Drop) oder XML" & @CRLF & _
                    "Autor: Florian Köppel / KI-Assistenz")
        EndSwitch
    WEnd
EndFunc

; =====================================================================================================
; DATEI ÖFFNEN
; =====================================================================================================
Func _OpenFile()
    Local $sFile = FileOpenDialog("ZUGFeRD-Datei auswählen (PDF oder XML)", @ScriptDir, "ZUGFeRD/Factur-X (*.pdf;*.xml)|PDF (*.pdf)|XML (*.xml)|Alle (*.*)", 1)
    If @error Then Return
    _LoadFile($sFile)
EndFunc

Func _LoadFile($sFile)
    If Not FileExists($sFile) Then
        GUICtrlSetData($g_hStatus, "❌ Datei nicht gefunden: " & $sFile)
        Return
    EndIf

    Local $sParse = $sFile
    Local $sInfo = "✅ " & $sFile
    $g_sOrigPdf = ""

    ; --- PDF? Dann eingebettete ZUGFeRD/Factur-X XML extrahieren ---
    If StringLower(StringRight($sFile, 4)) = ".pdf" Then
        GUICtrlSetData($g_hStatus, "📄 PDF erkannt – extrahiere Rechnungs-XML …")
        GUICtrlSetData($g_hEdit_Validation, "Extrahiere eingebettete XML aus der PDF …")
        Local $sXml = _ExtractXMLFromPDF($sFile)
        If $sXml = "" Then
            GUICtrlSetData($g_hStatus, "❌ Keine ZUGFeRD/Factur-X XML in der PDF gefunden")
            GUICtrlSetData($g_hEdit_Validation, _
                "In der PDF wurde keine eingebettete Rechnungs-XML gefunden." & @CRLF & @CRLF & _
                "Mögliche Ursachen:" & @CRLF & _
                "  - Die PDF ist keine ZUGFeRD/Factur-X-Rechnung (keine XML eingebettet)" & @CRLF & _
                "  - PowerShell ist nicht verfügbar (für komprimierte XML nötig)" & @CRLF & @CRLF & _
                "Tipp: Die XML auch direkt als .xml laden.")
            GUICtrlSetState($g_hBtnParse, $GUI_DISABLE)
            Return
        EndIf
        $sParse = $sXml
        $sInfo = "✅ " & $sFile & "   (XML aus PDF extrahiert)"
        $g_sOrigPdf = $sFile
    EndIf

    $g_sLastFile = $sParse
    GUICtrlSetData($g_hStatus, $sInfo)
    GUICtrlSetState($g_hBtnParse, $GUI_ENABLE)
    GUICtrlSetState($g_hBtnExport, $GUI_DISABLE)

    ; Roh-XML Vorschau (ersten 5000 Zeichen)
    Local $sRaw = FileRead($sParse)
    GUICtrlSetData($g_hEdit_XML, StringLeft($sRaw, 5000) & @CRLF & @CRLF & _
        "... [Datei ist " & Int(StringLen($sRaw)/1024) & " KB, gekürzt auf 5 KB]")
    ; Direkt parsen (kein extra Klick nötig)
    _ParseAndDisplay()
EndFunc

; =====================================================================================================
; PARSEN
; =====================================================================================================
Func _ParseAndDisplay()
    If $g_sLastFile = "" Or Not FileExists($g_sLastFile) Then
        MsgBox($MB_ICONWARNING, "Fehler", "Bitte zuerst eine XML-Datei öffnen.")
        Return
    EndIf

    ; XML-Dokument laden
    Local $oXML = ObjCreate("MSXML2.DOMDocument.6.0")
    If @error Then
        MsgBox($MB_ICONERROR, "Fehler", "MSXML2.DOMDocument.6.0 konnte nicht erstellt werden." & @CRLF & _
            "Bitte installieren Sie MSXML 6.0 oder neuer.")
        Return
    EndIf
    $oXML.async = False
    $oXML.validateOnParse = False  ; erstmal nur laden
    $oXML.preserveWhiteSpace = True
    $oXML.load($g_sLastFile)

    If @error Or $oXML.parseError.errorCode <> 0 Then
        Local $sErr = "❌ XML-Fehler: " & $oXML.parseError.reason & " (Zeile " & $oXML.parseError.line & ")"
        GUICtrlSetData($g_hEdit_Validation, $sErr)
        GUICtrlSetData($g_hStatus, "❌ Parse-Fehler")
        Return
    EndIf

    ; Format erkennen
    Local $sFormat = _DetectFormat($oXML)
    GUICtrlSetData($g_hStatus, "📄 " & $g_sLastFile & "  |  Format: " & $sFormat)

    ; Felder auslesen
    Local $aFields[1][2]
    Local $aValidation[1]

    _ClearListView($g_hListView)

    If StringInStr($sFormat, "ZUGFeRD 2") Or StringInStr($sFormat, "Factur-X") Then
        _ParseCII($oXML, $aFields, $aValidation)
    ElseIf StringInStr($sFormat, "ZUGFeRD 1") Then
        _ParseZUG1($oXML, $aFields, $aValidation)
    ElseIf StringInStr($sFormat, "XRechnung") Then
        _ParseXRechnung($oXML, $aFields, $aValidation)
    Else
        ; Fallback: Universal-Parser
        _ParseCII($oXML, $aFields, $aValidation)
    EndIf

    ; Felder in ListView anzeigen
    _PopulateListView($aFields)
    GUICtrlSetState($g_hBtnExport, $GUI_ENABLE)

    ; Validierung anzeigen
    _ShowValidation($aValidation, $sFormat)

    ; Summen anzeigen
    _ShowSums($oXML, $sFormat)
EndFunc

; =====================================================================================================
; FORMAT ERKENNEN
; =====================================================================================================
Func _DetectFormat(ByRef $oXML)
    ; Root-Element ohne Namensraum-Abhaengigkeit bestimmen
    Local $oRoot = $oXML.documentElement
    Local $sRoot = ""
    If IsObj($oRoot) Then $sRoot = $oRoot.nodeName
    Local $aSplit = StringSplit($sRoot, ":")
    If $aSplit[0] >= 2 Then $sRoot = $aSplit[2]

    ; 1) XRechnung / UBL  (Root 'Invoice' mit CustomizationID)
    Local $oCust = $oXML.selectSingleNode("//*[local-name()='CustomizationID']")
    If $sRoot = "Invoice" Then
        Local $sCust = ""
        If IsObj($oCust) Then $sCust = $oCust.text
        Return "XRechnung / UBL 2.x  (" & $sCust & ")"
    EndIf

    ; 2) ZUGFeRD 2.x / Factur-X (CII)
    If $sRoot = "CrossIndustryInvoice" Then
        Local $oProfile = $oXML.selectSingleNode("//*[local-name()='GuidelineSpecifiedDocumentContextParameter']/*[local-name()='ID']")
        If IsObj($oProfile) Then
            Local $sProf = $oProfile.text
            If StringInStr($sProf, "minimum")  Then Return "Factur-X / ZUGFeRD 2.x (MINIMUM)"
            If StringInStr($sProf, "basicwl")  Then Return "Factur-X / ZUGFeRD 2.x (BASIC WL)"
            If StringInStr($sProf, "basic")    Then Return "Factur-X / ZUGFeRD 2.x (BASIC)"
            If StringInStr($sProf, "comfort")  Then Return "Factur-X / ZUGFeRD 2.x (COMFORT)"
            If StringInStr($sProf, "extended") Then Return "Factur-X / ZUGFeRD 2.x (EXTENDED)"
            If StringInStr($sProf, "en16931")  Then Return "Factur-X / ZUGFeRD 2.x (EN16931)"
            Return "Factur-X / ZUGFeRD 2.x (" & $sProf & ")"
        EndIf
        Return "Factur-X / ZUGFeRD 2.x (CII)"
    EndIf

    ; 3) ZUGFeRD 1.x
    If $sRoot = "CrossIndustryDocument" Then Return "ZUGFeRD 1.x"

    ; 4) Unbekannt
    Return "Unbekannt (Root: " & $sRoot & ") - versuche CII-Parser"
EndFunc

; =====================================================================================================
; CII-PARSER (ZUGFeRD 2.x / Factur-X)
; =====================================================================================================
Func _ParseCII(ByRef $oXML, ByRef $aFields, ByRef $aValidation)
    ReDim $aValidation[100]
    Local $iV = 0

    ; --- Tester / Kontext ---
    Local $oTest = $oXML.selectSingleNode("//*[local-name()='ExchangedDocumentContext']/*[local-name()='TestIndicator']")
    If IsObj($oTest) Then
        Local $sTest = $oTest.text
        If $sTest = "true" Then
            _AddField($aFields, "Test-/Entwurf", "JA  ⚠️ Testdokument (keine echte Rechnung)")
            $aValidation[$iV] = "⚠️ HINWEIS: Das Dokument ist als TEST/ENTWURF markiert!"
            $iV += 1
        ElseIf $sTest = "false" Then
            _AddField($aFields, "Test-/Entwurf", "Nein (echte Rechnung)")
        Else
            _AddField($aFields, "Testindikator", $sTest)
        EndIf
    EndIf

    ; --- Profil ---
    Local $oProf = $oXML.selectSingleNode("//*[local-name()='GuidelineSpecifiedDocumentContextParameter']/*[local-name()='ID']")
    If IsObj($oProf) Then
        _AddField($aFields, "Profil", $oProf.text)
        $aValidation[$iV] = "Profil: " & $oProf.text
        $iV += 1
    EndIf

    ; --- ExchangedDocument (Kopfdaten) ---
    Local $oExDoc = $oXML.selectSingleNode("//*[local-name()='ExchangedDocument']")

    ; Rechnungsnummer
    Local $sInvNo = _GetText($oXML, "//*[local-name()='ExchangedDocument']/*[local-name()='ID']")
    If $sInvNo = "" Then $sInvNo = _GetText($oXML, "//*[local-name()='ID']")
    _AddField($aFields, "Rechnungsnummer", $sInvNo)
    If $sInvNo = "" Then
        $aValidation[$iV] = "❌ FEHLT: Rechnungsnummer (ID in ExchangedDocument)"
        $iV += 1
    Else
        $aValidation[$iV] = "✅ Rechnungsnummer: " & $sInvNo
        $iV += 1
    EndIf

    ; Rechnungstyp / TypeCode
    Local $sTypeCode = _GetText($oXML, "//*[local-name()='ExchangedDocument']/*[local-name()='TypeCode']")
    If $sTypeCode = "" Then $sTypeCode = _GetText($oXML, "//*[local-name()='TypeCode']")
    If $sTypeCode <> "" Then
        Local $sTypeDesc = _InvoiceTypeDesc($sTypeCode)
        _AddField($aFields, "Rechnungstyp", $sTypeCode & " (" & $sTypeDesc & ")")
    EndIf

    ; Rechnungsname
    Local $sDocName = _GetText($oXML, "//*[local-name()='ExchangedDocument']/*[local-name()='Name']")
    If $sDocName <> "" Then _AddField($aFields, "Dokument-Name", $sDocName)

    ; Datum
    Local $sIssueDate = _GetDate($oXML, "//*[local-name()='ExchangedDocument']/*[local-name()='IssueDateTime']")
    If $sIssueDate = "" Then $sIssueDate = _GetDateFromRAM($oXML, "//*[local-name()='IssueDateTime']")
    _AddField($aFields, "Rechnungsdatum", $sIssueDate)
    If $sIssueDate = "" Then
        $aValidation[$iV] = "❌ FEHLT: Rechnungsdatum (IssueDateTime)"
        $iV += 1
    Else
        $aValidation[$iV] = "✅ Rechnungsdatum: " & $sIssueDate
        $iV += 1
    EndIf

    ; Währung (aus Settlement)
    Local $sCurrency = _GetText($oXML, "//*[local-name()='InvoiceCurrencyCode']")
    If $sCurrency = "" Then
        ; XRechnung-Fallback
        $sCurrency = _GetText($oXML, "//*[local-name()='PaymentCurrencyCode']")
    EndIf
    If $sCurrency <> "" Then
        _AddField($aFields, "Währung", $sCurrency)
        If StringLen($sCurrency) <> 3 Then
            $aValidation[$iV] = "⚠️ Währungscode '" & $sCurrency & "' entspricht nicht ISO 4217 (3 Zeichen)"
            $iV += 1
        EndIf
    EndIf

    ; --- VERKÄUFER (SellerTradeParty) ---
    _ParseSeller($oXML, $aFields, $aValidation, $iV)

    ; --- KÄUFER (BuyerTradeParty) ---
    _ParseBuyer($oXML, $aFields, $aValidation, $iV)

    ; --- Zahlungsbedingungen / Fälligkeit ---
    Local $sDueDate = _GetDate($oXML, "//*[local-name()='SpecifiedTradeSettlementPaymentTerms']/*[local-name()='DueDateDateTime']")
    If $sDueDate = "" Then $sDueDate = _GetDateFromRAM($oXML, "//*[local-name()='DueDateDateTime']")
    _AddField($aFields, "Fälligkeitsdatum", $sDueDate)

    ; Zahlungsbedingung-Beschreibung
    Local $sPayTerms = _GetText($oXML, "//*[local-name()='SpecifiedTradePaymentTerms']/*[local-name()='Description']")
    If $sPayTerms <> "" Then _AddField($aFields, "Zahlungsbedingung", $sPayTerms)

    ; Lieferdatum
    Local $sDelivDate = _GetDate($oXML, "//*[local-name()='ActualDeliverySupplyChainEvent']/*[local-name()='OccurrenceDateTime']")
    If $sDelivDate <> "" Then _AddField($aFields, "Lieferdatum", $sDelivDate)

    ; Zahlungsweg / Bankverbindung (ZUGFeRD 2.x: SpecifiedTradeSettlementPaymentMeans)
    Local $sPayMeans = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementPaymentMeans']/*[local-name()='TypeCode']")
    Local $sIBAN = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementPaymentMeans']/*[local-name()='PayeePartyCreditorFinancialAccount']/*[local-name()='IBANID']")
    If $sIBAN = "" Then $sIBAN = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementPaymentMeans']/*[local-name()='PayerPartyDebtorFinancialAccount']/*[local-name()='IBANID']")
    If $sIBAN = "" Then $sIBAN = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementFinancialCard']/*[local-name()='ID']") ; Alt-Fallback
    Local $sBIC = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementPaymentMeans']/*[local-name()='PayeeSpecifiedCreditorFinancialInstitution']/*[local-name()='BICID']")
    Local $sAcctName = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementPaymentMeans']/*[local-name()='PayeePartyCreditorFinancialAccount']/*[local-name()='AccountName']")
    If $sPayMeans <> "" Then _AddField($aFields, "Zahlungsart (Code)", $sPayMeans)
    If $sIBAN <> "" Then _AddField($aFields, "IBAN", $sIBAN)
    If $sBIC <> "" Then _AddField($aFields, "BIC", $sBIC)
    If $sAcctName <> "" Then _AddField($aFields, "Kontoinhaber", $sAcctName)

    ; --- Summen (HeaderTradeSettlement) ---
    _ParseTotals($oXML, $aFields, $aValidation, $iV)

    ; --- Line Items ---
    _ParseLineItems($oXML, $aFields, $aValidation, $iV)

    ; Validation-Check: Netto + Steuer = Brutto
    _ValidateMath($aFields, $aValidation, $iV)

    ReDim $aValidation[$iV]
EndFunc

; =====================================================================================================
; VERKÄUFER
; =====================================================================================================
Func _ParseSeller(ByRef $oXML, ByRef $aFields, ByRef $aValidation, ByRef $iV)
    Local $sName = _GetText($oXML, "//*[local-name()='SellerTradeParty']/*[local-name()='Name']")
    If $sName = "" Then $sName = _GetText($oXML, "//*[local-name()='SellerTradeParty']/*[local-name()='SpecifiedOrganization']/*[local-name()='Name']")
    If $sName = "" Then $sName = _GetText($oXML, "(//*[local-name()='Name'])[1]") ; Fallback: erstes Name-Element

    If $sName <> "" Then
        _AddField($aFields, "Verkäufer (Name)", $sName)
        $aValidation[$iV] = "✅ Verkäufer: " & $sName
        $iV += 1
    Else
        $aValidation[$iV] = "❌ FEHLT: Verkäufer-Name (SellerTradeParty/Name)"
        $iV += 1
    EndIf

    Local $sStreet = _GetText($oXML, "//*[local-name()='SellerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='LineOne']")
    If $sStreet = "" Then $sStreet = _GetText($oXML, "//*[local-name()='SellerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='LineOne']")
    If $sStreet <> "" Then _AddField($aFields, "  Straße", $sStreet)

    Local $sZIP = _GetText($oXML, "//*[local-name()='SellerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='PostcodeCode']")
    If $sZIP = "" Then $sZIP = _GetText($oXML, "//*[local-name()='SellerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='PostcodeCode']")
    If $sZIP <> "" Then
        Local $sCity = _GetText($oXML, "//*[local-name()='SellerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='CityName']")
        Local $sCountry = _GetText($oXML, "//*[local-name()='SellerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='CountryID']")
        _AddField($aFields, "  PLZ / Ort / Land", $sZIP & " " & $sCity & " / " & $sCountry)
    EndIf

    Local $sVAT = _GetTaxRegID($oXML, "//*[local-name()='SellerTradeParty']/*[local-name()='SpecifiedTaxRegistration']", "VA")
    If $sVAT <> "" Then
        _AddField($aFields, "  USt-IdNr.", $sVAT)
    Else
        $aValidation[$iV] = "⚠️ Verkäufer-USt-IdNr. nicht gefunden (SchemeID/TypeCode='VA')"
        $iV += 1
    EndIf
EndFunc

; =====================================================================================================
; KÄUFER
; =====================================================================================================
Func _ParseBuyer(ByRef $oXML, ByRef $aFields, ByRef $aValidation, ByRef $iV)
    Local $sName = _GetText($oXML, "//*[local-name()='BuyerTradeParty']/*[local-name()='Name']")
    If $sName = "" Then $sName = _GetText($oXML, "//*[local-name()='BuyerTradeParty']/*[local-name()='SpecifiedOrganization']/*[local-name()='Name']")
    If $sName <> "" Then
        _AddField($aFields, "Käufer (Name)", $sName)
        $aValidation[$iV] = "✅ Käufer: " & $sName
        $iV += 1
    Else
        $aValidation[$iV] = "❌ FEHLT: Käufer-Name"
        $iV += 1
    EndIf

    Local $sStreet = _GetText($oXML, "//*[local-name()='BuyerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='LineOne']")
    If $sStreet <> "" Then _AddField($aFields, "  Straße", $sStreet)

    Local $sZIP = _GetText($oXML, "//*[local-name()='BuyerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='PostcodeCode']")
    If $sZIP <> "" Then
        Local $sCity = _GetText($oXML, "//*[local-name()='BuyerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='CityName']")
        Local $sCountry = _GetText($oXML, "//*[local-name()='BuyerTradeParty']/*[local-name()='PostalTradeAddress']/*[local-name()='CountryID']")
        _AddField($aFields, "  PLZ / Ort / Land", $sZIP & " " & $sCity & " / " & $sCountry)
    EndIf

    Local $sVAT = _GetTaxRegID($oXML, "//*[local-name()='BuyerTradeParty']/*[local-name()='SpecifiedTaxRegistration']", "VA")
    If $sVAT <> "" Then _AddField($aFields, "  USt-IdNr.", $sVAT)

    ; Steuernummer (FC)
    Local $sTaxNo = _GetTaxRegID($oXML, "//*[local-name()='BuyerTradeParty']/*[local-name()='SpecifiedTaxRegistration']", "FC")
    If $sTaxNo <> "" Then _AddField($aFields, "  Steuernummer", $sTaxNo)
EndFunc

; =====================================================================================================
; SUMMEN
; =====================================================================================================
Func _ParseTotals(ByRef $oXML, ByRef $aFields, ByRef $aValidation, ByRef $iV)
    ; Nettobetrag
    Local $sNet = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementHeaderMonetarySummation']/*[local-name()='LineTotalAmount']")
    If $sNet = "" Then $sNet = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementMonetarySummation']/*[local-name()='LineTotalAmount']")
    If $sNet <> "" Then _AddField($aFields, "Gesamt-Netto", $sNet & " " & _GetCurrencyHint($oXML))

    ; Steuer pro Satz (nur Header-Ebene, nicht Positions-Ebene)
    Local $oNodes = $oXML.selectNodes("//*[local-name()='ApplicableHeaderTradeSettlement']/*[local-name()='ApplicableTradeTax']")
    If Not IsObj($oNodes) Or $oNodes.length = 0 Then $oNodes = $oXML.selectNodes("//*[local-name()='ApplicableTradeTax']")
    If IsObj($oNodes) And $oNodes.length > 0 Then
        For $i = 0 To $oNodes.length - 1
            Local $oNode = $oNodes.item($i)
            Local $sRate = "??"
            Local $oRate = $oNode.selectSingleNode("*[local-name()='RateApplicablePercent']")
            If IsObj($oRate) Then $sRate = $oRate.text & "%"
            Local $sTaxAmt = ""
            Local $oTaxAmt = $oNode.selectSingleNode("*[local-name()='CalculatedAmount']")
            If IsObj($oTaxAmt) Then $sTaxAmt = $oTaxAmt.text
            Local $sCalcBase = ""
            Local $oBase = $oNode.selectSingleNode("*[local-name()='BasisAmount']")
            If IsObj($oBase) Then $sCalcBase = $oBase.text
            _AddField($aFields, "  MwSt. (" & $sRate & ")", $sTaxAmt & " (Bemessungsgrundlage: " & $sCalcBase & ")")
        Next
    EndIf

    ; Gesamtsteuer
    Local $sTaxTotal = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementHeaderMonetarySummation']/*[local-name()='TaxTotalAmount']")
    If $sTaxTotal = "" Then $sTaxTotal = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementMonetarySummation']/*[local-name()='TaxTotalAmount']")
    If $sTaxTotal <> "" Then _AddField($aFields, "Gesamt-Steuer", $sTaxTotal & " " & _GetCurrencyHint($oXML))

    ; Bruttobetrag
    Local $sGross = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementHeaderMonetarySummation']/*[local-name()='GrandTotalAmount']")
    If $sGross = "" Then $sGross = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementMonetarySummation']/*[local-name()='GrandTotalAmount']")
    If $sGross = "" Then $sGross = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementHeaderMonetarySummation']/*[local-name()='DuePayableAmount']")
    If $sGross = "" Then $sGross = _GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementMonetarySummation']/*[local-name()='DuePayableAmount']")
    _AddField($aFields, "Gesamt-Brutto (fällig)", $sGross & " " & _GetCurrencyHint($oXML))
    If $sGross = "" Then
        $aValidation[$iV] = "❌ FEHLT: Gesamtbetrag (GrandTotalAmount / DuePayableAmount)"
        $iV += 1
    EndIf
EndFunc

; =====================================================================================================
; LINE ITEMS
; =====================================================================================================
Func _ParseLineItems(ByRef $oXML, ByRef $aFields, ByRef $aValidation, ByRef $iV)
    Local $oItems = $oXML.selectNodes("//*[local-name()='IncludedSupplyChainTradeLineItem']")
    If Not IsObj($oItems) Or $oItems.length = 0 Then
        $oItems = $oXML.selectNodes("//*[local-name()='SupplyChainTradeLineItem']")
    EndIf
    If Not IsObj($oItems) Or $oItems.length = 0 Then Return

    _AddField($aFields, "Positionen (Anzahl)", $oItems.length)

    _GUICtrlListView_DeleteAllItems($g_hLineListview)

    For $i = 0 To $oItems.length - 1
        Local $oItem = $oItems.item($i)
        Local $iLineNo = $i + 1
        ; Positionsnummer (aus LineID, sonst fortlaufend)
        Local $sLineID = _GetTextFromNode($oItem, "*[local-name()='AssociatedDocumentLineDocument']/*[local-name()='LineID']")
        If $sLineID = "" Then $sLineID = String($iLineNo)

        ; Artikelnummer
        Local $sArtNo = _GetTextFromNode($oItem, "*[local-name()='SpecifiedTradeProduct']/*[local-name()='GlobalID']")
        If $sArtNo = "" Then $sArtNo = _GetTextFromNode($oItem, "*[local-name()='SpecifiedTradeProduct']/*[local-name()='SellerAssignedID']")
        If $sArtNo = "" Then $sArtNo = _GetTextFromNode($oItem, "*[local-name()='SpecifiedTradeProduct']/*[local-name()='BuyerAssignedID']")

        ; Bezeichnung
        Local $sDesc = _GetTextFromNode($oItem, "*[local-name()='SpecifiedTradeProduct']/*[local-name()='Name']")

        ; Menge + Einheit: SpecifiedLineTradeDelivery/*[local-name()='BilledQuantity']
        Local $sQty  = _GetTextFromNode($oItem, "*[local-name()='SpecifiedLineTradeDelivery']/*[local-name()='BilledQuantity']")
        Local $sUnit = ""
        Local $oQtyNode = $oItem.selectSingleNode("*[local-name()='SpecifiedLineTradeDelivery']/*[local-name()='BilledQuantity']")
        If IsObj($oQtyNode) Then $sUnit = $oQtyNode.getAttribute("unitCode")
        If $sUnit = "" Then $sUnit = "C62"
        Local $sUnitDesc = _UnitDesc($sUnit)

        ; Einzel-Nettopreis: SpecifiedLineTradeAgreement/NetPriceProductTradePrice/ChargeAmount
        Local $sUnitPrice = _GetTextFromNode($oItem, "*[local-name()='SpecifiedLineTradeAgreement']/*[local-name()='NetPriceProductTradePrice']/*[local-name()='ChargeAmount']")
        If $sUnitPrice = "" Then $sUnitPrice = _GetTextFromNode($oItem, "*[local-name()='NetPriceProductTradePrice']/*[local-name()='ChargeAmount']")

        ; Nettobetrag der Position: SpecifiedLineTradeSettlement/...LineMonetarySummation/LineTotalAmount
        Local $sNetAmt = _GetTextFromNode($oItem, "*[local-name()='SpecifiedLineTradeSettlement']/*[local-name()='SpecifiedTradeSettlementLineMonetarySummation']/*[local-name()='LineTotalAmount']")
        If $sNetAmt = "" Then $sNetAmt = _GetTextFromNode($oItem, "*[local-name()='SpecifiedTradeSettlementLineMonetarySummation']/*[local-name()='LineTotalAmount']")

        ; Steuersatz + Steuerbetrag der Position
        Local $sTaxRate = _GetTextFromNode($oItem, "*[local-name()='SpecifiedLineTradeSettlement']/*[local-name()='ApplicableTradeTax']/*[local-name()='RateApplicablePercent']")
        If $sTaxRate = "" Then $sTaxRate = _GetTextFromNode($oItem, "*[local-name()='ApplicableTradeTax']/*[local-name()='RateApplicablePercent']")
        Local $sTaxAmt = _GetTextFromNode($oItem, "*[local-name()='SpecifiedLineTradeSettlement']/*[local-name()='ApplicableTradeTax']/*[local-name()='CalculatedAmount']")

        _GUICtrlListView_AddItem($g_hLineListview, $sLineID)
        _GUICtrlListView_SetItemText($g_hLineListview, $i, $sLineID, 0)
        _GUICtrlListView_SetItemText($g_hLineListview, $i, $sArtNo, 1)
        _GUICtrlListView_SetItemText($g_hLineListview, $i, $sDesc, 2)
        _GUICtrlListView_SetItemText($g_hLineListview, $i, $sQty, 3)
        _GUICtrlListView_SetItemText($g_hLineListview, $i, $sUnitDesc, 4)
        _GUICtrlListView_SetItemText($g_hLineListview, $i, $sUnitPrice, 5)
        _GUICtrlListView_SetItemText($g_hLineListview, $i, $sNetAmt, 6)
        _GUICtrlListView_SetItemText($g_hLineListview, $i, _TrimNum($sTaxRate) & "%", 7)
        _GUICtrlListView_SetItemText($g_hLineListview, $i, $sTaxAmt, 8)
    Next

    $aValidation[$iV] = "✅ Positionen: " & $oItems.length & " Stück"
    $iV += 1
EndFunc

; =====================================================================================================
; SUMMEN-BLOCK (unten)
; =====================================================================================================
Func _ShowSums(ByRef $oXML, $sFormat)
    Local $sNet = _Money(_GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementHeaderMonetarySummation']/*[local-name()='LineTotalAmount']"))
    Local $sTax = _Money(_GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementHeaderMonetarySummation']/*[local-name()='TaxTotalAmount']"))
    Local $sGross = _Money(_GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementHeaderMonetarySummation']/*[local-name()='GrandTotalAmount']"))
    If $sGross = "" Then $sGross = _Money(_GetText($oXML, "//*[local-name()='SpecifiedTradeSettlementHeaderMonetarySummation']/*[local-name()='DuePayableAmount']"))
    Local $sCcy = _GetText($oXML, "//*[local-name()='InvoiceCurrencyCode']")
    If $sCcy = "" Then $sCcy = _GetText($oXML, "//*[local-name()='DocumentCurrencyCode']")
    Local $sSums = "Netto:            " & $sNet & " " & $sCcy & @CRLF
    $sSums &= "Steuer:           " & $sTax & " " & $sCcy & @CRLF
    $sSums &= "Brutto (fällig):  " & $sGross & " " & $sCcy
    GUICtrlSetData($g_hSumEdit, $sSums)
EndFunc

; =====================================================================================================
; VALIDIERUNG
; =====================================================================================================
Func _ShowValidation(ByRef $aValidation, $sFormat)
    Local $sText = "═══════════════════════════════════════" & @CRLF
    $sText &= "FORMAT: " & $sFormat & @CRLF
    $sText &= "═══════════════════════════════════════" & @CRLF & @CRLF

    Local $iOK = 0, $iFail = 0, $iWarn = 0
    For $i = 0 To UBound($aValidation) - 1
        If $aValidation[$i] = "" Then ContinueLoop
        $sText &= $aValidation[$i] & @CRLF
        If StringLeft($aValidation[$i], 1) = "✅" Then $iOK += 1
        If StringLeft($aValidation[$i], 1) = "❌" Then $iFail += 1
        If StringLeft($aValidation[$i], 1) = "⚠" Then $iWarn += 1
    Next

    $sText &= @CRLF & "═══════════════════════════════════════" & @CRLF
    $sText &= "✅ OK: " & $iOK & "   ❌ Fehler: " & $iFail & "   ⚠ Warnungen: " & $iWarn & @CRLF

    ; Validierung mit MSXML Schema
    $sText &= @CRLF & "═══════════════════════════════════════" & @CRLF
    $sText &= "SCHEMA-VALIDIERUNG (sofern Schema verfügbar):" & @CRLF

    ; XSD-Prüfung (lazy: validateOnParse = True nochmal laden)
    Local $oXML2 = ObjCreate("MSXML2.DOMDocument.6.0")
    $oXML2.async = False
    $oXML2.validateOnParse = True
    $oXML2.load($g_sLastFile)
    If $oXML2.parseError.errorCode <> 0 Then
        $sText &= "⚠ Schema-Warnung: " & $oXML2.parseError.reason & " (Zeile " & $oXML2.parseError.line & ")" & @CRLF
        $sText &= "  (Das XML ist dennoch wohlgeformt – möglicherweise fehlt ein Schema-Import)" & @CRLF
    Else
        $sText &= "✅ XML ist gültig (wohlgeformt + gegen geladene Schemata validiert)" & @CRLF
    EndIf

    GUICtrlSetData($g_hEdit_Validation, $sText)
EndFunc

; =====================================================================================================
; MATHEMATISCHE PRÜFUNG
; =====================================================================================================
Func _ValidateMath(ByRef $aFields, ByRef $aValidation, ByRef $iV)
    Local $sNet = _GetFieldValue($aFields, "Gesamt-Netto")
    Local $sTax = _GetFieldValue($aFields, "Gesamt-Steuer")
    Local $sGross = _GetFieldValue($aFields, "Gesamt-Brutto (fällig)")

    ; Nur prüfen wenn vorhanden
    If $sNet = "" Or $sTax = "" Or $sGross = "" Then Return

    ; Zahlen extrahieren (erste Zahl aus String rausholen)
    Local $fNet = _ExtractNumber($sNet)
    Local $fTax = _ExtractNumber($sTax)
    Local $fGross = _ExtractNumber($sGross)

    If $fNet = -1 Or $fTax = -1 Or $fGross = -1 Then Return

    Local $fCalcGross = $fNet + $fTax
    Local $fDiff = Abs($fCalcGross - $fGross)

    If $fDiff < 0.01 Then
        $aValidation[$iV] = "✅ Mathematische Prüfung: Netto + Steuer = Brutto  (" & _
            Round($fNet, 2) & " + " & Round($fTax, 2) & " = " & Round($fCalcGross, 2) & " ≈ " & Round($fGross, 2) & ")"
        $iV += 1
    Else
        $aValidation[$iV] = "⚠️ Mathematische Abweichung: Netto(" & Round($fNet,2) & ") + Steuer(" & Round($fTax,2) & _
            ") = " & Round($fCalcGross,2) & " ≠ Brutto(" & Round($fGross,2) & ")  (Diff: " & Round($fDiff,2) & ")"
        $iV += 1
    EndIf
EndFunc

; =====================================================================================================
; ZUGFeRD 1.x PARSER
; =====================================================================================================
Func _ParseZUG1(ByRef $oXML, ByRef $aFields, ByRef $aValidation)
    ReDim $aValidation[50]
    Local $iV = 0

    Local $sInvNo = _GetText($oXML, "//*[local-name()='Invoice']/*[local-name()='Header']/*[local-name()='InvoiceNumber']")
    If $sInvNo = "" Then $sInvNo = _GetText($oXML, "//*[local-name()='InvoiceNumber']")
    _AddField($aFields, "Rechnungsnummer", $sInvNo)
    $aValidation[$iV] = "✅ Rechnungsnummer: " & $sInvNo
    $iV += 1

    Local $sDate = _GetText($oXML, "//*[local-name()='Invoice']/*[local-name()='Header']/*[local-name()='InvoiceDate']")
    If $sDate = "" Then $sDate = _GetText($oXML, "//*[local-name()='InvoiceDate']")
    _AddField($aFields, "Rechnungsdatum", $sDate)

    Local $sCurr = _GetText($oXML, "//*[local-name()='Invoice']/*[local-name()='Header']/*[local-name()='DocumentCurrency']")
    If $sCurr = "" Then $sCurr = _GetText($oXML, "//*[local-name()='DocumentCurrency']")
    _AddField($aFields, "Währung", $sCurr)

    ; Verkäufer
    Local $sSeller = _GetText($oXML, "//*[local-name()='Invoice']/*[local-name()='Header']/*[local-name()='Seller']/*[local-name()='Name']")
    If $sSeller <> "" Then
        _AddField($aFields, "Verkäufer", $sSeller)
        $aValidation[$iV] = "✅ Verkäufer: " & $sSeller
        $iV += 1
    EndIf

    Local $sBuyer = _GetText($oXML, "//*[local-name()='Invoice']/*[local-name()='Header']/*[local-name()='Buyer']/*[local-name()='Name']")
    If $sBuyer <> "" Then
        _AddField($aFields, "Käufer", $sBuyer)
        $aValidation[$iV] = "✅ Käufer: " & $sBuyer
        $iV += 1
    EndIf

    ; Summen
    Local $sNet = _GetText($oXML, "//*[local-name()='Invoice']/*[local-name()='Header']/*[local-name()='Summation']/*[local-name()='NetAmount']")
    If $sNet = "" Then $sNet = _GetText($oXML, "//*[local-name()='NetAmount']")
    _AddField($aFields, "Gesamt-Netto", $sNet & " " & $sCurr)

    Local $sTax = _GetText($oXML, "//*[local-name()='Invoice']/*[local-name()='Header']/*[local-name()='Summation']/*[local-name()='TaxAmount']")
    If $sTax = "" Then $sTax = _GetText($oXML, "//*[local-name()='TaxAmount']")
    _AddField($aFields, "Gesamt-Steuer", $sTax & " " & $sCurr)

    Local $sGross = _GetText($oXML, "//*[local-name()='Invoice']/*[local-name()='Header']/*[local-name()='Summation']/*[local-name()='GrossAmount']")
    If $sGross = "" Then $sGross = _GetText($oXML, "//*[local-name()='GrossAmount']")
    If $sGross = "" Then $sGross = _GetText($oXML, "//*[local-name()='Invoice']/*[local-name()='Header']/*[local-name()='Summation']/*[local-name()='TotalAmount']")
    _AddField($aFields, "Gesamt-Brutto", $sGross & " " & $sCurr)

    ; Positionen
    Local $oItems = $oXML.selectNodes("//*[local-name()='Invoice']/*[local-name()='ItemList']/*[local-name()='Item']")
    If IsObj($oItems) Then _AddField($aFields, "Positionen", $oItems.length)

    ReDim $aValidation[$iV]
EndFunc

; =====================================================================================================
; XRECHNUNG PARSER (UBL)
; =====================================================================================================
Func _ParseXRechnung(ByRef $oXML, ByRef $aFields, ByRef $aValidation)
    ReDim $aValidation[50]
    Local $iV = 0

    Local $sInvNo = _GetText($oXML, "//*[local-name()='ID']")
    _AddField($aFields, "Rechnungsnummer", $sInvNo)
    $aValidation[$iV] = "✅ Rechnungsnummer: " & $sInvNo
    $iV += 1

    Local $sDate = _GetText($oXML, "//*[local-name()='IssueDate']")
    _AddField($aFields, "Rechnungsdatum", $sDate)

    Local $sCcy = _GetText($oXML, "//*[local-name()='DocumentCurrencyCode']")
    _AddField($aFields, "Währung", $sCcy)

    Local $sSeller = _GetText($oXML, "//*[local-name()='AccountingSupplierParty']/*[local-name()='Party']/*[local-name()='PartyName']/*[local-name()='Name']")
    If $sSeller = "" Then $sSeller = _GetText($oXML, "//*[local-name()='AccountingSupplierParty']/*[local-name()='Party']/*[local-name()='Name']")
    _AddField($aFields, "Verkäufer", $sSeller)
    $aValidation[$iV] = "✅ Verkäufer: " & $sSeller
    $iV += 1

    Local $sBuyer = _GetText($oXML, "//*[local-name()='AccountingCustomerParty']/*[local-name()='Party']/*[local-name()='PartyName']/*[local-name()='Name']")
    _AddField($aFields, "Käufer", $sBuyer)
    $aValidation[$iV] = "✅ Käufer: " & $sBuyer
    $iV += 1

    Local $sNet = _GetText($oXML, "//*[local-name()='LegalMonetaryTotal']/*[local-name()='LineExtensionAmount']")
    _AddField($aFields, "Gesamt-Netto", $sNet & " " & $sCcy)

    Local $sTax = _GetText($oXML, "//*[local-name()='LegalMonetaryTotal']/*[local-name()='TaxTotalAmount']")
    _AddField($aFields, "Gesamt-Steuer", $sTax & " " & $sCcy)

    Local $sGross = _GetText($oXML, "//*[local-name()='LegalMonetaryTotal']/*[local-name()='PayableAmount']")
    _AddField($aFields, "Gesamt-Brutto", $sGross & " " & $sCcy)

    ReDim $aValidation[$iV]
EndFunc

; =====================================================================================================
; HILFSFUNKTIONEN
; =====================================================================================================

; Steuer-Registrierung (USt-IdNr 'VA' / Steuernummer 'FC') aus SpecifiedTaxRegistration holen.
; Unterstuetzt BEIDE Varianten: <ram:ID schemeID="VA"> und <ram:TypeCode>VA</ram:TypeCode>.
Func _GetTaxRegID(ByRef $oXML, $sNodePath, $sScheme)
    Local $oNodes = $oXML.selectNodes($sNodePath)
    If Not IsObj($oNodes) Then Return ""
    For $i = 0 To $oNodes.length - 1
        Local $oNode = $oNodes.item($i)
        Local $oID = $oNode.selectSingleNode("*[local-name()='ID']")
        If Not IsObj($oID) Then ContinueLoop
        ; Variante 1: schemeID-Attribut
        If $oID.getAttribute("schemeID") = $sScheme Then Return StringStripWS($oID.text, 3)
        ; Variante 2: TypeCode-Kindelement
        Local $oType = $oNode.selectSingleNode("*[local-name()='TypeCode']")
        If IsObj($oType) And $oType.text = $sScheme Then Return StringStripWS($oID.text, 3)
    Next
    Return ""
EndFunc

; UN/ECE-Einheitencode (unitCode) -> lesbare Kurzbeschreibung
Func _UnitDesc($sCode)
    Switch StringUpper($sCode)
        Case "C62", "H87", "EA", "PCE", "NAR", "ZZ"
            Return "Stk"
        Case "KGM"
            Return "kg"
        Case "GRM"
            Return "g"
        Case "TNE"
            Return "t"
        Case "MTR"
            Return "m"
        Case "CMT"
            Return "cm"
        Case "MMT"
            Return "mm"
        Case "MTK"
            Return "m²"
        Case "MTQ"
            Return "m³"
        Case "LTR"
            Return "l"
        Case "MLT", "ML"
            Return "ml"
        Case "HUR"
            Return "Std"
        Case "MIN"
            Return "Min"
        Case "DAY"
            Return "Tag"
        Case "KMT"
            Return "km"
        Case "WEE"
            Return "Woche"
        Case "MON"
            Return "Monat"
        Case Else
            Return $sCode
    EndSwitch
EndFunc

; Zahl als String ohne überflüssige Nullen (19.000 -> 19, 39.90 -> 39.9)
Func _TrimNum($sVal)
    If StringStripWS($sVal, 3) = "" Then Return ""
    Return StringFormat("%g", Number(StringReplace($sVal, ",", ".")))
EndFunc

; Text aus Node holen (mit XPath relativ zum Document)
Func _GetText(ByRef $oXML, $sXPath)
    Local $oNode = $oXML.selectSingleNode($sXPath)
    If IsObj($oNode) Then Return StringStripWS($oNode.text, 3)
    Return ""
EndFunc

; Text aus relativem XPath (Kind-Ebene)
Func _GetTextFromNode(ByRef $oNode, $sXPath)
    Local $oChild = $oNode.selectSingleNode($sXPath)
    If IsObj($oChild) Then Return StringStripWS($oChild.text, 3)
    Return ""
EndFunc

; Datum aus CII (DateTimeString)
Func _GetDate(ByRef $oXML, $sXPath)
    Local $oNode = $oXML.selectSingleNode($sXPath)
    If IsObj($oNode) Then
        Local $oStr = $oNode.selectSingleNode("*[local-name()='DateTimeString']")
        If IsObj($oStr) Then Return StringStripWS($oStr.text, 3)
        Return StringStripWS($oNode.text, 3)
    EndIf
    Return ""
EndFunc

Func _GetDateFromRAM(ByRef $oXML, $sXPath)
    Local $oNode = $oXML.selectSingleNode($sXPath)
    If IsObj($oNode) Then
        Local $oStr = $oNode.selectSingleNode("*[local-name()='DateTimeString']")
        If IsObj($oStr) Then Return StringStripWS($oStr.text, 3)
        Return StringStripWS($oNode.text, 3)
    EndIf
    Return ""
EndFunc

; Attribut-Text
Func _GetAttributeText(ByRef $oXML, $sXPath)
    ; Not needed for simple cases, use selectSingleNode with attribute
    Local $oNode = $oXML.selectSingleNode($sXPath)
    If IsObj($oNode) Then Return StringStripWS($oNode.text, 3)
    Return ""
EndFunc

; Currency-Hinweis
Func _GetCurrencyHint(ByRef $oXML)
    Local $sCcy = _GetText($oXML, "//*[local-name()='InvoiceCurrencyCode']")
    If $sCcy = "" Then $sCcy = _GetText($oXML, "//*[local-name()='DocumentCurrencyCode']")
    If $sCcy = "" Then $sCcy = _GetText($oXML, "//*[local-name()='DocumentCurrency']")
    Return $sCcy
EndFunc

; Invoice Type Code Beschreibung
Func _InvoiceTypeDesc($sCode)
    Local $aMap[11][2] = [ _
        ["380", "Rechnung (Invoice)"], _
        ["381", "Gutschrift (Credit Note)"], _
        ["383", "Gutschrift"], _
        ["384", "Korrekturrechnung"], _
        ["385", "Korrekturgutschrift"], _
        ["386", "Vorauszahlungsrechnung"], _
        ["388", "Umsatzsteuer-Korrektur"], _
        ["389", "Selbst fakturierte Rechnung" ], _
        ["390", "Selbst fakturierte Gutschrift"], _
        ["393", "Rechnung mit Aufschlag"], _
        ["395", "Vorauszahlung" ]]
    For $i = 0 To UBound($aMap) - 1
        If $aMap[$i][0] = $sCode Then Return $aMap[$i][1]
    Next
    Return "Unbekannter Typ (" & $sCode & ")"
EndFunc

; Feld zum Array hinzufügen
Func _AddField(ByRef $aFields, $sName, $sValue)
    Local $iU = UBound($aFields)
    ReDim $aFields[$iU + 1][2]
    $aFields[$iU][0] = $sName
    $aFields[$iU][1] = $sValue
EndFunc

; Feldwert aus Array lesen
Func _GetFieldValue(ByRef $aFields, $sName)
    For $i = 0 To UBound($aFields) - 1
        If $aFields[$i][0] = $sName Then Return $aFields[$i][1]
    Next
    Return ""
EndFunc

; Zahl aus String extrahieren (z.B. "123.45 EUR" -> 123.45)
Func _ExtractNumber($sStr)
    Local $aMatches = StringRegExp($sStr, "([0-9]+[\.\,]?[0-9]*)", 1)
    If @error Then Return -1
    Local $sNum = StringReplace($aMatches[0], ",", ".")
    Return Number($sNum)
EndFunc

; ListView befüllen
Func _PopulateListView(ByRef $aFields)
    _GUICtrlListView_DeleteAllItems($g_hListView)
    For $i = 0 To UBound($aFields) - 1
        If $aFields[$i][0] = "" Then ContinueLoop
        ; Wichtig: den von AddItem zurueckgegebenen Zeilenindex verwenden
        ; (nicht $i - das Array-Element 0 wird uebersprungen, sonst Verschiebung)
        Local $iIdx = _GUICtrlListView_AddItem($g_hListView, $aFields[$i][0])
        _GUICtrlListView_AddSubItem($g_hListView, $iIdx, $aFields[$i][1], 1)
    Next
EndFunc

; ListView leeren
Func _ClearListView(ByRef $hListView)
    _GUICtrlListView_DeleteAllItems($hListView)
EndFunc

; =====================================================================================================
; EXPORT
; =====================================================================================================
Func _ExportToText()
    Local $sFile = FileSaveDialog("Speichern als...", @ScriptDir, "Text (*.txt)|JSON (*.json)|Alle (*.*)", 16, "ZUGFeRD_Ausgabe.txt")
    If @error Then Return

    Local $sText = "ZUGFeRD Reader Ausgabe" & @CRLF
    $sText &= "Datei: " & $g_sLastFile & @CRLF
    $sText &= "═══════════════════════════════════════" & @CRLF & @CRLF

    ; Felder aus ListView
    For $i = 0 To _GUICtrlListView_GetItemCount($g_hListView) - 1
        Local $sField = _GUICtrlListView_GetItemText($g_hListView, $i, 0)
        Local $sValue = _GUICtrlListView_GetItemText($g_hListView, $i, 1)
        $sText &= $sField & ": " & $sValue & @CRLF
    Next

    $sText &= @CRLF & "═══════════════════════════════════════" & @CRLF
    $sText &= "VALIDIERUNG:" & @CRLF
    $sText &= GUICtrlRead($g_hEdit_Validation) & @CRLF

    FileWrite($sFile, $sText)
    If @error Then
        MsgBox($MB_ICONERROR, "Fehler", "Export fehlgeschlagen.")
    Else
        MsgBox($MB_ICONINFORMATION, "Export", "Gespeichert:" & @CRLF & $sFile)
        ShellExecute($sFile)
    EndIf
EndFunc

; =====================================================================================================
; PDF: eingebettete ZUGFeRD/Factur-X XML extrahieren (via PowerShell/.NET)
; =====================================================================================================
Func _ExtractXMLFromPDF($sPdf)
    Local $sHelper = @TempDir & "\zg_pdf_extract.ps1"
    Local $sOut    = @TempDir & "\zg_extract.xml"
    If FileExists($sOut) Then FileDelete($sOut)

    ; PowerShell-Helfer schreiben (ANSI, ohne BOM)
    Local $hF = FileOpen($sHelper, 2)   ; 2 = überschreiben/erzeugen
    If $hF = -1 Then Return ""
    FileWrite($hF, $g_sPDFHelper)
    FileClose($hF)

    ; PowerShell ausführen
    Local $sPS = @SystemDir & "\WindowsPowerShell\v1.0\powershell.exe"
    If Not FileExists($sPS) Then $sPS = "powershell.exe"
    Local $sCmd = '"' & $sPS & '" -NoProfile -ExecutionPolicy Bypass -File "' & $sHelper & '" -Pdf "' & $sPdf & '" -Out "' & $sOut & '"'
    RunWait($sCmd, @TempDir, @SW_HIDE)

    If FileExists($sOut) Then
        If FileGetSize($sOut) > 200 Then Return $sOut
    EndIf
    Return ""
EndFunc

; =====================================================================================================
; DRAG & DROP (Datei ins Fenster ziehen) - reine DllCall-Variante, keine WinAPIEx-UDFs nötig
; =====================================================================================================
Func WM_DROPFILES($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg, $lParam
    Local $aCnt = DllCall("shell32.dll", "uint", "DragQueryFileW", "hwnd", $wParam, "uint", 0xFFFFFFFF, "ptr", 0, "uint", 0)
    If Not @error And IsArray($aCnt) And $aCnt[0] > 0 Then
        Local $tBuf = DllStructCreate("wchar[32768]")
        DllCall("shell32.dll", "uint", "DragQueryFileW", "hwnd", $wParam, "uint", 0, "ptr", DllStructGetPtr($tBuf), "uint", 32768)
        Local $sFile = DllStructGetData($tBuf, 1)
        If $sFile <> "" Then _LoadFile($sFile)
    EndIf
    DllCall("shell32.dll", "none", "DragFinish", "hwnd", $wParam)
    Return 0
EndFunc


; =====================================================================================================
; SICHTKOPIE  -  lesbare Ansicht der geparsten Daten + optionales Anhängen an die PDF
; =====================================================================================================
Func _LVField($sKey)
    For $i = 0 To _GUICtrlListView_GetItemCount($g_hListView) - 1
        If StringStripWS(_GUICtrlListView_GetItemText($g_hListView, $i, 0), 3) = $sKey Then Return _GUICtrlListView_GetItemText($g_hListView, $i, 1)
    Next
    Return ""
EndFunc

Func _FileName($sPath)
    Local $i = StringInStr($sPath, "\", 0, -1)
    If $i > 0 Then Return StringMid($sPath, $i + 1)
    Return $sPath
EndFunc

Func _Money($s)
    Local $a = StringRegExp($s, "[0-9]+(?:\.[0-9]+)?", 1)
    If @error Then Return StringStripWS($s, 3)
    Return StringReplace(StringFormat("%.2f", Number($a[0])), ".", ",")
EndFunc

Func _QtyNum($s)
    Local $a = StringRegExp($s, "[0-9]+(?:\.[0-9]+)?", 1)
    If @error Then Return StringStripWS($s, 3)
    Return StringReplace(StringFormat("%g", Number($a[0])), ".", ",")
EndFunc

Func _Rate($s)
    Return _QtyNum($s)
EndFunc

Func _DeDate($s)
    Local $t = StringStripWS($s, 3)
    If StringRegExp($t, "^[0-9]{8}$") Then Return StringMid($t, 7, 2) & "." & StringMid($t, 5, 2) & "." & StringLeft($t, 4)
    Return $t
EndFunc

Func _AddLine(ByRef $aLines, ByRef $n, $sText)
    ReDim $aLines[$n + 1]
    $aLines[$n] = $sText
    $n += 1
EndFunc

Func _SichtLines()
    Local $aLines[1], $n = 0
    Local $sSrc = $g_sLastFile
    If $g_sOrigPdf <> "" Then $sSrc = $g_sOrigPdf
    Local $sCur = _LVField("Währung")
    If $sCur = "" Then $sCur = "EUR"

    _AddLine($aLines, $n, "TITLE" & @TAB & "ZUGFeRD / Factur-X Sichtkopie" & @TAB & "Quelle: " & _FileName($sSrc))

    For $i = 0 To _GUICtrlListView_GetItemCount($g_hListView) - 1
        Local $sF = StringStripWS(_GUICtrlListView_GetItemText($g_hListView, $i, 0), 3)
        Local $sV = _GUICtrlListView_GetItemText($g_hListView, $i, 1)
        If $sF = "Gesamt-Netto" Or $sF = "Gesamt-Steuer" Or $sF = "Gesamt-Brutto (fällig)" Then ContinueLoop
        If StringInStr($sF, "datum") Then $sV = _DeDate($sV)
        _AddLine($aLines, $n, "META" & @TAB & $sF & @TAB & $sV)
    Next

    Local $hc = _GUICtrlListView_GetItemCount($g_hLineListview)
    If $hc > 0 Then
        _AddLine($aLines, $n, "HEAD" & @TAB & "POSITIONEN")
        _AddLine($aLines, $n, "TH" & @TAB & "Pos" & @TAB & "Art-Nr" & @TAB & "Bezeichnung" & @TAB & "Menge" & @TAB & "Einzelpreis" & @TAB & "Netto" & @TAB & "St%" & @TAB & "St")
        For $i = 0 To $hc - 1
            Local $r = "TR"
            $r &= @TAB & _GUICtrlListView_GetItemText($g_hLineListview, $i, 0)
            $r &= @TAB & _GUICtrlListView_GetItemText($g_hLineListview, $i, 1)
            $r &= @TAB & _GUICtrlListView_GetItemText($g_hLineListview, $i, 2)
            $r &= @TAB & _QtyNum(_GUICtrlListView_GetItemText($g_hLineListview, $i, 3)) & " " & _GUICtrlListView_GetItemText($g_hLineListview, $i, 4)
            $r &= @TAB & _Money(_GUICtrlListView_GetItemText($g_hLineListview, $i, 5))
            $r &= @TAB & _Money(_GUICtrlListView_GetItemText($g_hLineListview, $i, 6))
            $r &= @TAB & _Rate(_GUICtrlListView_GetItemText($g_hLineListview, $i, 7))
            $r &= @TAB & _Money(_GUICtrlListView_GetItemText($g_hLineListview, $i, 8))
            _AddLine($aLines, $n, $r)
        Next
    EndIf

    Local $sNet = _LVField("Gesamt-Netto")
    Local $sTax = _LVField("Gesamt-Steuer")
    Local $sGross = _LVField("Gesamt-Brutto (fällig)")
    _AddLine($aLines, $n, "HEAD" & @TAB & "SUMMEN")
    _AddLine($aLines, $n, "TOT" & @TAB & "Netto" & @TAB & _Money($sNet) & " " & $sCur)
    _AddLine($aLines, $n, "TOT" & @TAB & "Steuer" & @TAB & _Money($sTax) & " " & $sCur)
    _AddLine($aLines, $n, "TOTB" & @TAB & "Brutto (fällig)" & @TAB & _Money($sGross) & " " & $sCur)

    Local $aV = StringSplit(StringReplace(GUICtrlRead($g_hEdit_Validation), @CRLF, @LF), @LF)
    For $i = 1 To $aV[0]
        Local $t = StringStripWS($aV[$i], 3)
        If $t <> "" And (StringInStr($t, "Mathematische") Or StringInStr($t, "OK:")) Then
            _AddLine($aLines, $n, "NOTE" & @TAB & $t)
        EndIf
    Next
    Return $aLines
EndFunc

Func _StripExt($sPath)
    Local $iDot = StringInStr($sPath, ".", 0, -1)
    If $iDot > 1 Then Return StringLeft($sPath, $iDot - 1)
    Return $sPath
EndFunc

Func _WritePSFile($sPath, $sContent)
    Local $h = FileOpen($sPath, 2)
    If $h = -1 Then Return False
    FileWrite($h, $sContent)
    FileClose($h)
    Return True
EndFunc

Func _MakeSichtkopie()
    If $g_sLastFile = "" Then
        MsgBox($MB_ICONWARNING, "Sichtkopie", "Bitte zuerst eine Datei laden.")
        Return
    EndIf
    If _GUICtrlListView_GetItemCount($g_hListView) = 0 Then
        MsgBox($MB_ICONWARNING, "Sichtkopie", "Keine geparsten Daten vorhanden. Bitte erst 'Parsen'.")
        Return
    EndIf

    Local $sDef
    If $g_sOrigPdf <> "" Then
        $sDef = _StripExt($g_sOrigPdf) & "_Sichtkopie.pdf"
    Else
        $sDef = _StripExt($g_sLastFile) & "_Sichtkopie.pdf"
    EndIf
    Local $sOut = FileSaveDialog("Sichtkopie speichern", @ScriptDir, "PDF (*.pdf)", 16, $sDef)
    If @error Then Return
    If StringLower(StringRight($sOut, 4)) <> ".pdf" Then $sOut &= ".pdf"
    ; Original niemals überschreiben
    If $g_sOrigPdf <> "" And StringLower($sOut) = StringLower($g_sOrigPdf) Then
        $sOut = _StripExt($sOut) & "_mitSichtkopie.pdf"
    EndIf

    ; Zeilen als UTF-8 ablegen
    Local $sLinesFile = @TempDir & "\zg_sicht_lines.txt"
    Local $aLines = _SichtLines()
    Local $h = FileOpen($sLinesFile, 2 + 128)
    If $h = -1 Then
        MsgBox($MB_ICONERROR, "Sichtkopie", "Temp-Datei konnte nicht geschrieben werden.")
        Return
    EndIf
    For $i = 0 To UBound($aLines) - 1
        FileWriteLine($h, $aLines[$i])
    Next
    FileClose($h)

    Local $sPSFile = @TempDir & "\zg_pdf_sicht.ps1"
    _WritePSFile($sPSFile, $g_sSichtHelper)

    Local $sPw = @SystemDir & "\WindowsPowerShell\v1.0\powershell.exe"
    If Not FileExists($sPw) Then $sPw = "powershell.exe"
    Local $sCmd = '"' & $sPw & '" -NoProfile -ExecutionPolicy Bypass -File "' & $sPSFile & _
        '" -Base "' & $g_sOrigPdf & '" -Lines "' & $sLinesFile & '" -Out "' & $sOut & '"'
    RunWait($sCmd, @TempDir, @SW_HIDE)

    If FileExists($sOut) And FileGetSize($sOut) > 200 Then
        Local $sWhat = "Sichtkopie erstellt"
        If $g_sOrigPdf <> "" Then $sWhat = "Sichtkopie an die Original-PDF angehaengt"
        MsgBox($MB_ICONINFORMATION, "Sichtkopie", $sWhat & ":" & @CRLF & $sOut)
        ShellExecute($sOut)
    Else
        MsgBox($MB_ICONERROR, "Sichtkopie", "Erstellung fehlgeschlagen." & @CRLF & "Pruefe, ob PowerShell verfuegbar ist.")
    EndIf
EndFunc
