$ErrorActionPreference = 'Stop'
$srcTxt = "apps/cuaderno-westminster/fuente/Confesion de fe Westminster - 1647 1903.txt"
$base = "apps/cuaderno-westminster/index.html"
$out = "apps/cuaderno-westminster/index.html"

$raw = [System.IO.File]::ReadAllText($srcTxt, [System.Text.Encoding]::UTF8)
$raw = $raw -replace "`r`n", "`n"
$lines = $raw -split "`n"

# find chapter headers
$hits = @()
for ($i = 0; $i -lt $lines.Length; $i++) {
  if ($lines[$i] -match '^CAP.{1,2}TULO\s+([IVXLCDM]+)\s*(.*)$') {
    $hits += [pscustomobject]@{ Idx = $i; Num = $Matches[1].Trim(); Title = ($Matches[2].Trim()) }
  }
}

function FixTitle([string]$num, [string]$t) {
  $t = ($t -replace '\s+', ' ').Trim()
  $map = @{
    'I'='De las Sagradas Escrituras'; 'II'='De Dios y de la Santísima Trinidad'; 'III'='Del Decreto Eterno de Dios'
    'IV'='De la Creación'; 'V'='De la Providencia'; 'VI'='De la Caída del Hombre, del Pecado y de su Castigo'
    'VII'='Del Pacto de Dios con el Hombre'; 'VIII'='De Cristo el Mediador'; 'IX'='Del Libre Albedrío'
    'X'='Del Llamamiento Eficaz'; 'XI'='De la Justificación'; 'XII'='De la Adopción'; 'XIII'='De la Santificación'
    'XIV'='De la Fe Salvadora'; 'XV'='Del Arrepentimiento para Vida'; 'XVI'='De las Buenas Obras'
    'XVII'='De la Perseverancia de los Santos'; 'XVIII'='De la Certeza de la Gracia y de la Salvación'
    'XIX'='De la Ley de Dios'; 'XX'='De la Libertad Cristiana y de la Libertad de Conciencia'
    'XXI'='Del Culto Religioso y del Día de Reposo'; 'XXII'='De los Juramentos Lícitos y de los Votos'
    'XXIII'='Del Magistrado Civil'; 'XXIV'='Del Matrimonio y del Divorcio'; 'XXV'='De la Iglesia'
    'XXVI'='De la Comunión de los Santos'; 'XXVII'='De los Sacramentos'; 'XXVIII'='Del Bautismo'
    'XXIX'='De la Cena del Señor'; 'XXX'='De las Censuras Eclesiásticas'; 'XXXI'='De los Sínodos y Concilios'
    'XXXII'='Del Estado del Hombre después de la Muerte y de la Resurrección de los Muertos'
    'XXXIII'='Del Juicio Final'; 'XXXIV'='Del Espíritu Santo'; 'XXXV'='Del Amor de Dios y de las Misiones'
  }
  if ($map.ContainsKey($num)) { return $map[$num] }
  return $t
}

$chapters = @()
for ($j = 0; $j -lt $hits.Count; $j++) {
  $start = $hits[$j].Idx + 1
  $end = if ($j -lt $hits.Count - 1) { $hits[$j + 1].Idx - 1 } else { $lines.Length - 1 }
  $num = $hits[$j].Num; $title = $hits[$j].Title
  if ($title -match 'LAS CENSURAS DE LA IGLESIA' -and $num -eq 'XXI') { $num = 'XXX' }
  if ($num -eq 'XXXII' -and $title -match 'Y LA\s*$') {
    $next = if ($start -le $end) { $lines[$start].Trim() } else { '' }
    if ($next -match '^RESURRECCION') { $title = ($title + ' ' + $next).Trim(); $start++ }
  }
  $title = FixTitle $num $title
  $bodyLines = @(); for ($k = $start; $k -le $end; $k++) { $bodyLines += $lines[$k] }
  $bi = 0; $bj = $bodyLines.Count - 1
  while ($bi -le $bj -and $bodyLines[$bi].Trim() -eq '') { $bi++ }
  while ($bj -ge $bi -and $bodyLines[$bj].Trim() -eq '') { $bj-- }
  if ($bj -ge $bi) { $bodyLines = $bodyLines[$bi..$bj] } else { $bodyLines = @() }
  $chapters += [pscustomobject]@{ n = $num; titulo = $title; lines = $bodyLines }
}

function CleanText([string]$s) {
  if ($null -eq $s) { return '' }
  $t = $s
  $t = ($t -replace '\s+', ' ').Trim()
  return $t
}

function ParseSections([string[]]$ls) {
  $sections = @()
  $curN = ''; $cur = New-Object System.Collections.Generic.List[string]
  function FlushSec {
    if ($curN -ne '' -or $cur.Count -gt 0) {
      $txt = CleanText(($cur -join ' '))
      if ($txt -ne '') { $script:sections += [pscustomobject]@{ n = $curN; texto = $txt } }
      $cur.Clear() | Out-Null
    }
  }
  $script:sections = @()
  foreach ($ln in $ls) {
    $t = $ln.TrimEnd()
    if ($t.Trim() -eq '') { continue }
    if ($t.Trim() -match '^([IVXLCDM]+)\.\s*$') { FlushSec; $curN = $Matches[1]; continue }
    $m = [regex]::Match($t.Trim(), '^([IVXLCDM]+)\.\s+(.*)$')
    if ($m.Success) { FlushSec; $curN = $m.Groups[1].Value; $cur.Add($m.Groups[2].Value.Trim()) | Out-Null; continue }
    $cur.Add($t.Trim()) | Out-Null
  }
  FlushSec
  # if no roman sections found, whole thing as one
  if ($script:sections.Count -eq 0) {
    $txt = CleanText(($ls -join ' '))
    if ($txt -ne '') { $script:sections = @([pscustomobject]@{ n = '1'; texto = $txt }) }
  }
  return $script:sections
}

function Paras([string[]]$ls) {
  $paras = @(); $cur = New-Object System.Collections.Generic.List[string]
  foreach ($ln in $ls) {
    $t = $ln.TrimEnd()
    if ($t.Trim() -eq '') { if ($cur.Count -gt 0) { $paras += (CleanText(($cur -join ' '))); $cur.Clear() | Out-Null } ; continue }
    # treat standalone roman marker as break
    if ($t.Trim() -match '^[IVXLCDM]+\.$') { if ($cur.Count -gt 0) { $paras += (CleanText(($cur -join ' '))); $cur.Clear() | Out-Null }; $cur.Add($t.Trim()) | Out-Null; continue }
    $cur.Add($t.Trim()) | Out-Null
  }
  if ($cur.Count -gt 0) { $paras += (CleanText(($cur -join ' '))) }
  $paras = $paras | Where-Object { $_ -and $_.Trim() -ne '' }
  return $paras
}

$lecciones = @()
$idx = 1
foreach ($c in $chapters) {
  $secs = ParseSections $c.lines
  $arts = @()
  foreach ($s in $secs) { $arts += [pscustomobject]@{ n = $s.n; texto = $s.texto } }
  $desarrollo = Paras $c.lines
  if ($desarrollo.Count -eq 0) { $desarrollo = @("Texto del capítulo pendiente.") }
  $lecciones += [pscustomobject]@{
    num = $idx
    cap = $c.n
    titulo = $c.titulo
    arts = @("Capítulo $($c.n)")
    articulos = $arts
    biblia = @([pscustomobject]@{ ref = "Referencias bíblicas del capítulo"; texto = "Este capítulo cita y desarrolla pasajes de las Sagradas Escrituras según la Confesión de Fe de Westminster (1647/1903). Revisa las referencias incluidas en el texto."; nota = "Las citas bíblicas aparecen integradas en las secciones del capítulo." })
    introduccion = "Este capítulo de la Confesión de Fe de Westminster trata sobre «$($c.titulo)». Lee con atención cada sección, observa su argumento y las referencias bíblicas que lo sostienen."
    desarrollo = $desarrollo
    aplicacion = @(
      "¿Cuál es la enseñanza principal de este capítulo y cómo se resume en sus secciones?",
      "¿Qué pasajes bíblicos se citan o se presuponen, y cómo fundamentan la doctrina expuesta?",
      "¿Qué implicaciones prácticas tiene este capítulo para la fe, la vida cristiana y la iglesia?"
    )
    cierre = "Señor, ilumina nuestro entendimiento por tu Espíritu para recibir tu verdad con fe, humildad y obediencia. Amén."
  }
  $idx++
}

$datos = [pscustomobject]@{
  datos = [pscustomobject]@{
    iglesia = "Iglesia Presbiteriana de Chile"
    denominacion = "Confesión de Fe de Westminster (1647/1903)"
    curso = "Cuaderno de Estudio"
    tema = "Cuaderno de Estudio: Confesión de Fe de Westminster"
    anio = "2026"
    lema = "Toda la Escritura es inspirada por Dios y útil para enseñar, para reprender, para corregir, para instruir en justicia."
    fundamentoGeneral = "2 Timoteo 3.16"
  }
  objetivoGeneral = "Facilitar el estudio ordenado de la Confesión de Fe de Westminster (1647/1903), capítulo por capítulo, promoviendo la lectura atenta del texto, la reflexión bíblica y el registro personal de aprendizaje."
  objetivosEspecificos = @(
    "Leer y comprender el contenido de cada capítulo de la Confesión de Fe de Westminster en su estructura original por secciones.",
    "Identificar los temas doctrinales principales y las referencias bíblicas que sustentan cada capítulo.",
    "Registrar notas, reflexiones y avances de estudio de manera personal y progresiva.",
    "Utilizar el cuaderno como apoyo para la enseñanza, el estudio individual o grupal y la memorización de la doctrina reformada."
  )
  modulos = @([pscustomobject]@{
    numero = 1
    libro = "Confesión de Fe"
    titulo = "Confesión de Fe de Westminster"
    articulos = "Capítulos I al $($chapters[$chapters.Count-1].n)"
    fundamento = [pscustomobject]@{
      ref = "Salmo 119.105"
      texto = "Lámpara es a mis pies tu palabra, y lumbrera a mi camino."
      nota = "El estudio de la confesión busca conducirnos a la Escritura como norma suprema de fe y vida."
    }
    objetivoGeneral = "Estudiar de forma sistemática los capítulos de la Confesión de Fe de Westminster, comprendiendo su doctrina y registrando el avance personal."
    lecciones = $lecciones
  })
}

$json = $datos | ConvertTo-Json -Depth 12 -Compress
$json = $json -replace '</script>', '<\/script>'

# Load base and transform
$html = [System.IO.File]::ReadAllText($base, [System.Text.Encoding]::UTF8)

# Replace DATOS
$html = [regex]::Replace($html, 'var LOGO_VERTICAL = LOGO_CRUZ;const DATOS = .*?;\s*(?=// ---- helpers)', "var LOGO_VERTICAL = LOGO_CRUZ;`nconst DATOS = $json;`n`n", [System.Text.RegularExpressions.RegexOptions]::Singleline)

# storage keys
$html = $html -replace "const KEY = 'ebd-libro-notas-v1';", "const KEY = 'cuaderno-westminster-notas-v1';"
$html = $html -replace "const PROG = 'ebd-libro-progreso-v1';", "const PROG = 'cuaderno-westminster-progreso-v1';"

# titles/meta
$html = $html -replace '<title>.*?</title>', '<title>Cuaderno de Estudio · Confesión de Fe de Westminster</title>'
$html = $html -replace '<meta name="description" content="[^"]*">', '<meta name="description" content="Cuaderno de estudio interactivo de la Confesión de Fe de Westminster (1647/1903).">'
$html = $html -replace '<strong style="font-size:13px;flex:1">.*?</strong>', '<strong style="font-size:13px;flex:1">Cuaderno de Estudio</strong>'
$html = $html -replace '<div class="an">EBD &middot; Adultos &middot; 2026</div>', '<div class="an">Confesión de Westminster · 1647/1903</div>'
$html = $html -replace '<div class="pagefoot">.*?</div>', '<div class="pagefoot"><b>Confesión de Fe de Westminster</b> &nbsp;&middot;&nbsp; Cuaderno de estudio &nbsp;&middot;&nbsp; 1647/1903</div>'

# portada and general labels inside render
$html = $html -replace 'Propositos del programa', 'Propósitos del cuaderno'
$html = $html -replace 'Propositos del Programa', 'Propósitos del cuaderno'
$html = $html -replace 'Educaci&oacute;n Cristiana &middot; Escuela B&#237;blica Dominical', 'Cuaderno de Estudio &middot; Libro Interactivo'
$html = $html -replace 'Clase de Adultos &middot; ', ''
$html = $html -replace '<div class="pr2"><span class="rol">Profesor</span><span class="nom">Pr\. Dagoberto Pe&ntilde;aloza A\.</span></div>', '<div class="pr2"><span class="rol">Documento</span><span class="nom">Confesión de Fe de Westminster</span></div>'

# "Como usar este libro" section adjustments
$html = $html -replace '<li><b>Articulos del Estatuto\.</b> El texto normativo de la leccion, citado literalmente\.</li>', '<li><b>Texto de la Confesión.</b> El contenido del capítulo, citado por secciones.</li>'
$html = $html -replace '<li><b>Fundamento biblico\.</b> Los pasajes que sostienen el tema\. El versiculo en azul es el ancla de la leccion\.</li>', '<li><b>Referencias bíblicas.</b> Indicaciones sobre las citas y fundamentos bíblicos presentes en el capítulo.</li>'
$html = $html -replace '<li><b>Introduccion y Desarrollo\.</b> La explicacion de por que el articulo dice lo que dice, y que significa para la iglesia\.</li>', '<li><b>Introducción y Desarrollo.</b> Orientación para la lectura y el texto completo del capítulo para su estudio.</li>'
$html = $html -replace '<li><b>Aplicacion\.</b> De la teoria a la vida de la congregacion\.</li>', '<li><b>Aplicación.</b> Preguntas para reflexionar y aplicar lo estudiado.</li>'
$html = $html -replace '<li>Los recuadros <span style="color:#8a6b22;font-weight:700">amarillos</span> marcan articulos cuyo texto debes verificar directamente en el Estatuto vigente: el programa original los menciona pero no los reproduce\.</li>', ''
$html = $html -replace 'Leccion como leida', 'capítulo como leído'
$html = $html -replace 'la leccion como leida', 'el capítulo como leído'

# dynamic totals / labels in render
$html = $html -replace "h \+= '<div class=`"prog`"><div class=`"pl`">Leccion '\+l\.num\+' de 15</div>';", "h += '<div class=`"prog`"><div class=`"pl`">Capítulo '+l.num+' de '+D.modulos[0].lecciones.length+'</div>';"
$html = $html -replace '<span class="n">Leccion ', '<span class="n">Capítulo '
$html = $html -replace '<h3 class="sub">Articulos del Estatuto</h3>', '<h3 class="sub">Texto de la Confesión</h3>'
$html = $html -replace '<h3 class="sub">Fundamento biblico</h3>', '<h3 class="sub">Referencias bíblicas</h3>'
$html = $html -replace '<h3 class="sub">Introduccion</h3>', '<h3 class="sub">Introducción</h3>'
$html = $html -replace '<h3 class="sub">Aplicacion</h3>', '<h3 class="sub">Aplicación</h3>'
$html = $html -replace '<h3 class="sub">Oracion final</h3>', '<h3 class="sub">Oración final</h3>'
$html = $html -replace 'Notas de estudio &middot; Leccion ', 'Notas de estudio &middot; Capítulo '
$html = $html -replace 'placeholder="Escribe aqui lo que quieras recordar de esta leccion\.\.\."', 'placeholder="Escribe aquí lo que quieras recordar de este capítulo..."'

# remove estatuto-specific warning logic
$html = $html -replace "const falta = a\.nota && a\.nota\.indexOf\('verificar'\)>-1 \|\| \/\^24-27\|.*?\.test\(String\(a\.n\)\);\s*", ""
$html = $html -replace "if \(falta\) h \+= '<div class=`"falt`"><b>Nota:</b> '\+esc\(a\.nota\|\|'El programa original no reproduce el texto de este articulo\. Verificalo en el Estatuto vigente antes de ense&ntilde;arlo\.'\)\+'</div>';\s*else ", ""

# cierre section text
$html = $html -replace "Sean glorificados el nombre de Dios, la santidad de su gracia y la unidad de su Esp&iacute;ritu en esta casa y en toda la '\+esc\(D\.datos\.denominacion\)\+'\. Am&eacute;n\.", "A Dios sea la gloria por su Palabra y su verdad. Amén."
$html = $html -replace "Que el Se&ntilde;or que comenz&oacute; a esta obra en Su iglesia laiak\$ hasta el fin\.", "Que el Señor nos conceda comprender y vivir conforme a las Sagradas Escrituras."

# progress total dynamic
$html = $html -replace "for \(let n=1; n<=15; n\+\+\)\{", "const totalLecciones = (DATOS.modulos[0] && DATOS.modulos[0].lecciones) ? DATOS.modulos[0].lecciones.length : 0;`n  for (let n=1; n<=totalLecciones; n++){"
$html = $html -replace "let d=0; for\(let n=1;n<=15;n\+\+\) if\(prog\[n\]\) d\+\+;", "let d=0; const totalLecciones=(DATOS.modulos[0]&&DATOS.modulos[0].lecciones)?DATOS.modulos[0].lecciones.length:0; for(let n=1;n<=totalLecciones;n++) if(prog[n]) d++;"
$html = $html -replace "if \(b\) b\.style\.width = \(d/15\*100\)\+'%';", "if (b) b.style.width = (totalLecciones? (d/totalLecciones*100):0)+'%';"
$html = $html -replace "if \(t\) t\.textContent = d\+' de 15 lecciones leidas';", "if (t) t.textContent = d+' de '+totalLecciones+' capítulos leídos';"

# toc sub-link labels
$html = $html -replace "\['art','Articulos del Estatuto'\]", "['art','Texto de la Confesión']"
$html = $html -replace "\['bib','Fundamento biblico'\]", "['bib','Referencias bíblicas']"
$html = $html -replace "\['intro','Introduccion'\]", "['intro','Introducción']"
$html = $html -replace "\['apl','Aplicacion'\]", "['apl','Aplicación']"
$html = $html -replace "\['orar','Oracion final'\]", "['orar','Oración final']"

# button texts
$html = $html -replace "Marcar como leida", "Marcar como leído"
$html = $html -replace "Leida", "Leído"

# general text fixes
$html = $html -replace 'se guarda en este navegador\.', 'se guarda en este navegador.'

[System.IO.File]::WriteAllText($out, $html, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "OK lecciones=$($lecciones.Count)"