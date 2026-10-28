# Genera Plantilla_Suficiencia.xlsx con Excel (requiere Microsoft Excel instalado en Windows).
# Uso:  powershell -ExecutionPolicy Bypass -File herramientas\generar_plantilla.ps1 [-Salida ruta\archivo.xlsx]
# Funciona con Excel en español o en inglés: los formatos y nombres se traducen al idioma de Excel.
param([string]$Salida = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Plantilla_Suficiencia.xlsx'))
$ErrorActionPreference = 'Stop'
$out = [System.IO.Path]::GetFullPath($Salida)

function RGB([int]$r, [int]$g, [int]$b) { return [int]($r + 256 * $g + 65536 * $b) }
$NAVY = RGB 31 56 100; $WHITE = RGB 255 255 255; $INFILL = RGB 221 235 247; $INFONT = RGB 0 51 153
$GRAY = RGB 242 242 242; $HDR = RGB 214 224 240; $BORDER = RGB 191 191 191
$REDF = RGB 253 226 225; $REDT = RGB 192 0 0; $GRNF = RGB 226 239 218; $GRNT = RGB 0 97 0
$AMBF = RGB 255 242 204; $AMBT = RGB 156 101 0
$M = [Type]::Missing

function SetV($ws, $a, $v) {
    if ($v -is [int] -or $v -is [double]) { $ws.Range([string]$a).Formula = ([double]$v).ToString([System.Globalization.CultureInfo]::InvariantCulture) }
    else { $ws.Range([string]$a).Value2 = [string]$v }
}
function SetF($ws, $a, $f) { $ws.Range($a).Formula = $f }
function Inp($ws, $a) { $r = $ws.Range($a); $r.Interior.Color = $INFILL; $r.Font.Color = $INFONT; $r.Locked = $false; $r.Borders.LineStyle = 1; $r.Borders.Color = $BORDER }
function Cal($ws, $a) { $r = $ws.Range($a); $r.Interior.Color = $GRAY; $r.Borders.LineStyle = 1; $r.Borders.Color = $BORDER }
function Hdr($ws, $a) { $r = $ws.Range($a); $r.Interior.Color = $HDR; $r.Font.Bold = $true; $r.WrapText = $true; $r.VerticalAlignment = -4108; $r.Borders.LineStyle = 1; $r.Borders.Color = $BORDER }
function Band($ws, $row, $last, $text) {
    $r = $ws.Range("A${row}:${last}${row}"); $r.Interior.Color = $NAVY; $r.Font.Color = $WHITE; $r.Font.Bold = $true
    $r.Font.Size = 11; $r.RowHeight = 22; $r.VerticalAlignment = -4108; $ws.Range("A$row").Value2 = $text
}
# Desde PowerShell, Excel lee los formatos con la configuración regional del equipo.
# Por eso los códigos se escriben con los separadores locales (se calculan al abrir Excel)
# y la fecha usa el formato corto de fecha que Excel aplica solo.
$script:SEP_DEC = '.'
$script:SEP_MIL = ','
$script:FMT_FECHA = $null
function FormatoLocal($f) { return $f.Replace(',', [string][char]1).Replace('.', $script:SEP_DEC).Replace([string][char]1, $script:SEP_MIL) }
function Fmt($ws, $a, $f) {
    if ($f -eq 'dd/mm/yyyy') { if ($script:FMT_FECHA) { $ws.Range($a).NumberFormatLocal = $script:FMT_FECHA }; return }
    $ws.Range($a).NumberFormatLocal = FormatoLocal $f
}
# Traduce una fórmula en inglés al idioma de Excel usando una celda auxiliar.
function TraducirFormula($ws, $formula) {
    $c = $ws.Range('XFD1')
    $c.Formula = $formula
    $local = ([string]$c.FormulaLocal) -replace '^=@', '='
    $c.Clear() | Out-Null
    return $local
}
function CFText($ws, $a, $txt, $fill, $font) { $fc = $ws.Range($a).FormatConditions.Add(9, $M, $M, $M, $txt, 0); $fc.Interior.Color = $fill; $fc.Font.Color = $font; $fc.Font.Bold = $true }
function CFAlerts($ws, $a) { CFText $ws $a ([string][char]0x26A0) $REDF $REDT; CFText $ws $a ([string][char]0x2714) $GRNF $GRNT }
function CFSemaforo($ws, $a) { CFText $ws $a 'VERDE' $GRNF $GRNT; CFText $ws $a 'META CUMPLIDA' $GRNF $GRNT; CFText $ws $a 'AMARILLO' $AMBF $AMBT; CFText $ws $a 'ROJO' $REDF $REDT }
function CFEtiqueta($ws, $a) { CFText $ws $a 'REEMPLAZAR' $REDF $REDT; CFText $ws $a 'EN RIESGO' $AMBF $AMBT; CFText $ws $a 'APORTA' $GRNF $GRNT }
function DVList($ws, $a, $src) { $v = $ws.Range($a).Validation; $v.Delete(); $v.Add(3, 1, 1, $src) | Out-Null; $v.IgnoreBlank = $true; $v.InCellDropdown = $true }
function DVPct($ws, $a) { $v = $ws.Range($a).Validation; $v.Delete(); $v.Add(2, 1, 1, '0', '1') | Out-Null; $v.ErrorTitle = 'Porcentaje'; $v.ErrorMessage = 'Escribe un porcentaje entre 0 % y 100 % (ej. 40 %).' }
function DVWhole($ws, $a) { $v = $ws.Range($a).Validation; $v.Delete(); $v.Add(1, 1, 7, '0') | Out-Null; $v.ErrorTitle = 'Casos'; $v.ErrorMessage = 'Escribe un número entero mayor o igual a 0.' }
function Widths($ws, $h) { foreach ($k in $h.Keys) { $ws.Range("${k}:${k}").ColumnWidth = $h[$k] } }
function NoGrid($x, $ws, $zoom) { try { $ws.Activate(); $x.ActiveWindow.DisplayGridlines = $false; $x.ActiveWindow.Zoom = $zoom } catch {} }
function OADate($s) { return [datetime]::ParseExact($s, 'yyyy-MM-dd', $null).ToOADate() }


# ---------------------------------------------------------------- indicator sheet
function New-HojaIndicador($x, $ws, $d, $banner, $isExample) {
    $ws.Cells.Font.Name = 'Arial'; $ws.Cells.Font.Size = 10
    Widths $ws @{A = 8; B = 46; C = 16; D = 16; E = 16; F = 16; G = 16; H = 16; I = 14; J = 13; K = 16; L = 13; M = 42; N = 24 }
    $ws.Cells.VerticalAlignment = -4108

    SetF $ws 'A1' '="Suficiencia del indicador: "&IF(C5="","(sin nombre)",C5)'
    $ws.Range('A1').Font.Size = 16; $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Color = $NAVY
    SetV $ws 'A2' $banner
    if ($isExample) { $ws.Range('A2:N2').Interior.Color = $AMBF; $ws.Range('A2').Font.Color = $AMBT; $ws.Range('A2').Font.Bold = $true }
    SetV $ws 'A3' 'Azul = lo escribes tú  ·  Gris = se calcula solo, no lo toques  ·  Rojo / ⚠ = alerta, revisa'
    $ws.Range('A3').Font.Italic = $true

    # PASO 1
    Band $ws 4 'N' 'PASO 1 · ¿Cuánto me falta?  (datos del indicador)'
    $lab = @{5 = 'Nombre del indicador'; 6 = 'Sentido'; 7 = 'Unidad'; 8 = 'Ejecutado actual'; 9 = 'Meta 2027'; 10 = 'Valor ideal (opcional)'; 11 = 'Problema total hoy  (el pastel)'; 12 = 'Brecha a la meta  (lo que debes recuperar)'; 13 = 'Total de casos analizados' }
    foreach ($k in $lab.Keys) { SetV $ws "B$k" $lab[$k] }
    $ws.Range('B5:B13').Font.Bold = $true
    Inp $ws 'C5:C10'; Inp $ws 'C13'; Cal $ws 'C11:C12'
    DVList $ws 'C6' '=Parametros!$B$19:$B$20'; DVWhole $ws 'C13'
    Fmt $ws 'C8:C12' '0.00'; Fmt $ws 'C13' '#,##0'
    SetF $ws 'D8:D10' '=IF($C$7="","",$C$7)'; SetF $ws 'D11:D12' '=$C$130'
    SetV $ws 'E5' 'Escribe el nombre completo del indicador.'
    SetF $ws 'E6' '=IF(C6="","⚠ Elige una opción de la lista.","¿Es mejor que el indicador suba o que baje?")'
    SetV $ws 'E7' 'Ej.: %, horas, puntos NPS.'
    SetF $ws 'E8' '=IF(AND(C7="%",N(C8)>0,N(C8)<1,N(C9)<1),"⚠ ¿Escribiste el símbolo %? Si tu valor es 2,8 %, escribe solo 2,8.","Valor real de hoy. Escribe solo el número, sin el símbolo %.")'
    SetV $ws 'E9' 'Valor que debes alcanzar en 2027.'
    SetV $ws 'E10' 'El mejor valor posible. Si lo dejas vacío: 100 si sube es mejor, 0 si baja es mejor.'
    SetV $ws 'E11' 'Distancia entre hoy y el ideal. Es TODO el problema que explican los dolores.'
    SetV $ws 'E12' 'Distancia entre hoy y la meta 2027. Es lo que las iniciativas deben recuperar.'
    SetV $ws 'E13' 'Ej.: clientes que se dieron de baja en el periodo. Sirve para calcular la prevalencia.'
    SetF $ws 'C11' '=IF(C8="",0,ABS(C131-C8))'
    SetF $ws 'C12' '=IF(OR(C6="",C8="",C9=""),0,MAX(0,IF(C6="Sube es mejor",C9-C8,C8-C9)))'
    $ws.Range('C11:C12').Font.Bold = $true
    CFAlerts $ws 'E6:E8'

    # PASO 2
    Band $ws 15 'N' 'PASO 2 · ¿Qué causa el problema?  (dolores = porciones del pastel)'
    SetV $ws 'B16' 'Dato del dolor'
    for ($i = 0; $i -lt 6; $i++) { $c = [char](67 + $i); SetV $ws "${c}16" ("Dolor " + ($i + 1)) }
    Hdr $ws 'B16:H16'
    $l2 = @{17 = 'Nombre del dolor'; 18 = 'Casos con este dolor'; 19 = '% Prevalencia  (casos ÷ total analizado)'; 20 = '% Atribución automática  (su parte del pastel)'; 21 = '% Atribución manual  (opcional)'; 22 = '% Atribución usada'; 23 = 'Evidencia / fuente del dato'; 24 = 'Revisión' }
    foreach ($k in $l2.Keys) { SetV $ws "B$k" $l2[$k] }
    $ws.Range('B17:B24').Font.Bold = $true
    Inp $ws 'C17:H18'; Inp $ws 'C21:H21'; Inp $ws 'C23:H23'; Cal $ws 'C19:H20'; Cal $ws 'C22:H22'
    $ws.Range('C17:H17').WrapText = $true; $ws.Range('C23:H23').WrapText = $true
    $ws.Range('C22:H22').Font.Bold = $true
    DVWhole $ws 'C18:H18'; DVPct $ws 'C21:H21'
    Fmt $ws 'C18:H18' '#,##0'; Fmt $ws 'C19:H19' '0%'; Fmt $ws 'C20:H22' '0.0%'
    SetF $ws 'C19:H19' '=IF(OR(N(C18)=0,N($C$13)=0),0,C18/$C$13)'
    SetF $ws 'C20:H20' '=IF(SUM($C$19:$H$19)=0,0,C19/SUM($C$19:$H$19))'
    SetF $ws 'C22:H22' '=IF(C21="",C20,C21)'
    SetF $ws 'C24' '=IF(SUM(C22:H22)=0,"⚠ Aún no hay dolores con casos.",IF(ABS(SUM(C22:H22)-1)<0.0005,"✔ La atribución suma 100 %.","⚠ La atribución suma "&FIXED(SUM(C22:H22)*100,1)&" %; debe sumar 100 %. Revisa la fila 21."))'
    SetF $ws 'C25' '="La suma de prevalencias es "&FIXED(SUM(C19:H19)*100,0)&" %. Puede pasar de 100 % porque un caso puede tener varios dolores. Es normal."'
    $ws.Range('C25').Font.Italic = $true
    CFAlerts $ws 'C24'
    $ws.Range('A17').RowHeight = 42; $ws.Range('A23').RowHeight = 56; $ws.Range('A16').RowHeight = 20

    # PASO 3
    Band $ws 27 'N' 'PASO 3 · ¿Qué vas a hacer?  En cada dolor escribe qué % resuelve la iniciativa si se hace completa (vacío = 0 %)'
    SetV $ws 'A28' 'N°'; SetV $ws 'B28' 'Iniciativa'
    SetF $ws 'C28:H28' '=IF(C17="","Dolor "&(COLUMN()-2),C17)'
    SetV $ws 'I28' '% Ejecución esperada a 2027'; SetV $ws 'J28' 'Estado'; SetV $ws 'K28' 'Responsable'
    SetV $ws 'L28' 'Fecha comprometida'; SetV $ws 'M28' 'Evidencia que respalda los %'; SetV $ws 'N28' 'Revisión'
    Hdr $ws 'A28:N28'; $ws.Range('A28').RowHeight = 44
    for ($j = 0; $j -lt 8; $j++) { SetV $ws ("A" + (29 + $j)) ("I" + ($j + 1)) }
    $ws.Range('A29:A36').Font.Bold = $true; $ws.Range('A29:A36').HorizontalAlignment = -4108
    Inp $ws 'B29:M36'; Cal $ws 'N29:N36'
    $ws.Range('B29:B36').WrapText = $true; $ws.Range('M29:M36').WrapText = $true
    $ws.Range('A29:A36').RowHeight = 40
    DVPct $ws 'C29:I36'; DVList $ws 'J29:J36' '=Parametros!$C$19:$C$21'
    Fmt $ws 'C29:I36' '0%'; Fmt $ws 'L29:L36' 'dd/mm/yyyy'
    $ws.Range('C29:L36').HorizontalAlignment = -4108
    SetF $ws 'N29:N36' '=IF(B29="","",IF(I29="","⚠ Falta % de ejecución",IF(M29="","⚠ Sin respaldo","✔ OK")))'
    SetV $ws 'B37' 'Revisión'; $ws.Range('B37').Font.Bold = $true
    SetF $ws 'C37' '=IF(COUNTIF(B29:B36,"?*")=0,"⚠ Aún no hay iniciativas.",IF(COUNTIF(N29:N36,"⚠*")>0,"⚠ Revisa la columna Revisión (N): faltan datos.","✔ Iniciativas completas."))'
    CFAlerts $ws 'N29:N36'; CFAlerts $ws 'C37'

    # PASO 4
    Band $ws 39 'N' 'PASO 4 · ¿Me alcanza?  (resultado · aquí no escribes nada)'
    SetV $ws 'B40' 'Resultado'; SetV $ws 'C40' 'Plan (todo al 100 %)'; SetV $ws 'D40' 'Realista (con tu % de ejecución)'
    Hdr $ws 'B40:D40'; $ws.Range('A40').RowHeight = 32
    $l4 = @{41 = 'Puntos recuperados'; 42 = 'Suficiencia  (recuperado ÷ brecha)'; 43 = 'Semáforo'; 44 = 'Valor proyectado 2027'; 45 = 'Brecha que aún faltaría'; 46 = 'Dolor principal'; 47 = 'En palabras simples' }
    foreach ($k in $l4.Keys) { SetV $ws "B$k" $l4[$k] }
    $ws.Range('B41:B47').Font.Bold = $true
    Cal $ws 'C41:D45'
    SetF $ws 'C41' '=SUM(C53:H53)'; SetF $ws 'D41' '=SUM(C54:H54)'
    SetF $ws 'C42:D42' '=IF($C$12=0,0,C41/$C$12)'
    SetF $ws 'C43:D43' '=IF($C$8="","SIN DATOS",IF($C$12=0,"META CUMPLIDA",LOOKUP(C42,Semaforo_Desde,Semaforo_Color)))'
    SetF $ws 'C44:D44' '=IF($C$8="","",IF($C$6="Sube es mejor",$C$8+C41,$C$8-C41))'
    SetF $ws 'C45:D45' '=MAX(0,$C$12-C41)'
    SetF $ws 'E41' '=$C$130'; SetF $ws 'E44' '=IF($C$7="","",$C$7)'; SetF $ws 'E45' '=$C$130'
    Fmt $ws 'C41:D41' '0.00'; Fmt $ws 'C42:D42' '0%'; Fmt $ws 'C44:D45' '0.00'
    $ws.Range('C42:D43').Font.Bold = $true; $ws.Range('C42:D42').Font.Size = 14; $ws.Range('C41:D45').HorizontalAlignment = -4108
    SetF $ws 'C46' '=IF($C$132=0,"",$C$133&" · explica el "&FIXED(INDEX(C22:H22,$C$132)*100,0)&" % del problema · resuelto (realista): "&FIXED(INDEX(C52:H52,$C$132)*100,0)&" %")'
    SetF $ws 'C47' '=IF($C$8="","Completa el Paso 1 para ver el resultado.",IF($C$12=0,"La meta ya se cumple hoy: no hay brecha que cerrar.","Si todas las iniciativas se cumplen al 100 %, cubres el "&FIXED(C42*100,0)&" % de la brecha. Con la ejecución realista cubres el "&FIXED(D42*100,0)&" %."))'
    $ws.Range('C47').Font.Bold = $true
    CFSemaforo $ws 'C43:D43'

    # Resultado por dolor
    SetV $ws 'B49' 'Resultado por dolor'
    SetF $ws 'C49:H49' '=IF(C17="","Dolor "&(COLUMN()-2),C17)'
    Hdr $ws 'B49:H49'; $ws.Range('A49').RowHeight = 42
    $l5 = @{50 = '% Atribución usada'; 51 = '% del dolor resuelto · Plan'; 52 = '% del dolor resuelto · Realista'; 53 = 'Puntos recuperados · Plan'; 54 = 'Puntos recuperados · Realista'; 55 = 'Alerta' }
    foreach ($k in $l5.Keys) { SetV $ws "B$k" $l5[$k] }
    $ws.Range('B50:B55').Font.Bold = $true
    Cal $ws 'C50:H55'
    SetF $ws 'C50:H50' '=C22'; SetF $ws 'C51:H51' '=C89'; SetF $ws 'C52:H52' '=C101'
    SetF $ws 'C53:H53' '=$C$11*C50*C51'; SetF $ws 'C54:H54' '=$C$11*C50*C52'
    SetF $ws 'C55:H55' '=IF(C17="","",IF(AND(C50>Cobertura_Atrib_Min,C52<Cobertura_Resuelto_Min),"⚠ SIN COBERTURA",IF(C52>=Cobertura_Resuelto_Min,"✔ Cubierto","○ Peso bajo")))'
    Fmt $ws 'C50:H52' '0%'; Fmt $ws 'C53:H54' '0.00'
    $ws.Range('C50:H55').HorizontalAlignment = -4108
    CFAlerts $ws 'C55:H55'

    # Aporte por iniciativa + ranking
    foreach ($hr in 57, 67) {
        if ($hr -eq 57) { SetV $ws "A$hr" 'N°' } else { SetV $ws "A$hr" 'Puesto' }
        SetV $ws "B$hr" $(if ($hr -eq 57) { '¿Cuánto aporta cada iniciativa? (realista)' } else { 'Ranking: de mayor a menor aporte' })
        SetV $ws "C$hr" 'Aporte (puntos)'; SetV $ws "D$hr" 'Aporte (% de la brecha)'; SetV $ws "E$hr" 'Suficiencia si falla'
        SetV $ws "F$hr" 'Etiqueta'; SetV $ws "G$hr" 'En palabras simples'
        Hdr $ws "A${hr}:G${hr}"; $ws.Range("A$hr").RowHeight = 32
    }
    for ($j = 0; $j -lt 8; $j++) { SetV $ws ("A" + (58 + $j)) ("I" + ($j + 1)); SetV $ws ("A" + (68 + $j)) ($j + 1) }
    $ws.Range('A58:A65').Font.Bold = $true; $ws.Range('A68:A75').Font.Bold = $true
    $ws.Range('A58:A75').HorizontalAlignment = -4108
    Cal $ws 'B58:F65'; Cal $ws 'B68:F75'
    SetF $ws 'B58:B65' '=IF(B29="","",B29)'
    SetF $ws 'C58:C65' '=IF($B29="","",K104)'
    SetF $ws 'D58:D65' '=IF($B29="","",L104)'
    SetF $ws 'E58:E65' '=IF($B29="","",J104)'
    SetF $ws 'F58:F65' '=IF($B29="","",IF(OR(N($I29)=0,L104<Umbral_Reemplazar),"REEMPLAZAR",IF(OR($J29="Retrasada",$J29="En riesgo"),"EN RIESGO","APORTA")))'
    SetF $ws 'G58:G65' '=IF($B29="","","Aporta "&FIXED(K104,2)&" "&$C$130&" ("&FIXED(L104*100,0)&" % de la brecha). Si falla, la suficiencia baja de "&FIXED($D$42*100,0)&" % a "&FIXED(J104*100,0)&" %.")'
    SetF $ws 'B68:G75' '=INDEX(B$58:B$65,$C114)'
    foreach ($rg in 'C58:C65', 'C68:C75') { Fmt $ws $rg '0.00' }
    foreach ($rg in 'D58:E65', 'D68:E75') { Fmt $ws $rg '0%' }
    $ws.Range('C58:F75').HorizontalAlignment = -4108
    CFEtiqueta $ws 'F58:F65'; CFEtiqueta $ws 'F68:F75'

    # CÁLCULOS
    Band $ws 77 'N' 'CÁLCULOS · no tocar · así se llega al resultado (puedes revisarlos paso a paso)'
    $blocks = @(
        @{h = 79; s = 80; t = 'A · Lo que DEJA cada iniciativa en cada dolor · PLAN   = 1 − efectividad'; f = '=1-MIN(Tope_Efectividad,N(C29))' },
        @{h = 91; s = 92; t = 'B · Lo que DEJA cada iniciativa en cada dolor · REALISTA   = 1 − efectividad × ejecución'; f = '=1-MIN(Tope_Efectividad,N(C29)*N($I29))' }
    )
    foreach ($b in $blocks) {
        $h = $b.h; $s = $b.s; $e = $s + 7
        SetV $ws "B$h" $b.t; SetF $ws "C${h}:H${h}" '=IF(C$17="","Dolor "&(COLUMN()-2),C$17)'; Hdr $ws "A${h}:H${h}"; $ws.Range("A$h").RowHeight = 42
        for ($j = 0; $j -lt 8; $j++) { SetV $ws ("A" + ($s + $j)) ("I" + ($j + 1)) }
        SetF $ws "B${s}:B${e}" '=IF(B29="","",B29)'
        SetF $ws "C${s}:H${e}" $b.f
        Cal $ws "B${s}:H$($e+2)"; Fmt $ws "C${s}:H$($e+2)" '0%'
    }
    SetV $ws 'B88' 'Lo que queda del dolor  (se multiplica la columna)'; SetF $ws 'C88:H88' '=PRODUCT(C80:C87)'
    SetV $ws 'B89' '% del dolor resuelto · Plan   = 1 − lo que queda'; SetF $ws 'C89:H89' '=1-C88'
    SetV $ws 'B100' 'Lo que queda del dolor  (se multiplica la columna)'; SetF $ws 'C100:H100' '=PRODUCT(C92:C99)'
    SetV $ws 'B101' '% del dolor resuelto · Realista   = 1 − lo que queda'; SetF $ws 'C101:H101' '=1-C100'
    $ws.Range('B88:H89').Font.Bold = $true; $ws.Range('B100:H101').Font.Bold = $true

    SetV $ws 'B103' 'C · Si la iniciativa falla: % del dolor resuelto SIN ella  (se divide para deshacer su multiplicación)'
    SetF $ws 'C103:H103' '=IF(C$17="","Dolor "&(COLUMN()-2),C$17)'
    SetV $ws 'I103' 'Puntos recuperados sin ella'; SetV $ws 'J103' 'Suficiencia sin ella'; SetV $ws 'K103' 'Aporte (puntos)'
    SetV $ws 'L103' 'Aporte (% de la brecha)'; SetV $ws 'M103' 'Orden para el ranking (auxiliar)'
    Hdr $ws 'A103:M103'; $ws.Range('A103').RowHeight = 42
    for ($j = 0; $j -lt 8; $j++) { SetV $ws ("A" + (104 + $j)) ("I" + ($j + 1)) }
    SetF $ws 'B104:B111' '=IF(B29="","",B29)'
    SetF $ws 'C104:H111' '=1-C$100/C92'
    SetF $ws 'I104:I111' '=$C$11*SUMPRODUCT($C$22:$H$22,C104:H104)'
    SetF $ws 'J104:J111' '=IF($C$12=0,0,I104/$C$12)'
    SetF $ws 'K104:K111' '=$D$41-I104'
    SetF $ws 'L104:L111' '=IF($C$12=0,0,K104/$C$12)'
    SetF $ws 'M104:M111' '=IF(B29="",-1,K104)-ROW()/1000000000'
    Cal $ws 'B104:M111'
    Fmt $ws 'C104:H111' '0%'; Fmt $ws 'I104:I111' '0.00'; Fmt $ws 'J104:J111' '0%'; Fmt $ws 'K104:K111' '0.000'; Fmt $ws 'L104:L111' '0%'; Fmt $ws 'M104:M111' '0.000000'

    SetV $ws 'A113' 'Puesto'; SetV $ws 'B113' 'D · Posición en el ranking'; SetV $ws 'C113' 'Fila de la iniciativa'
    Hdr $ws 'A113:C113'
    for ($j = 0; $j -lt 8; $j++) { SetV $ws ("A" + (114 + $j)) ($j + 1) }
    SetF $ws 'C114:C121' '=MATCH(LARGE($M$104:$M$111,A114),$M$104:$M$111,0)'
    Cal $ws 'C114:C121'

    SetV $ws 'B123' 'E · Datos del gráfico'; SetV $ws 'C123' 'Valor'; Hdr $ws 'B123:C123'
    SetV $ws 'B124' 'Ejecutado hoy'; SetV $ws 'B125' 'Proyectado realista'; SetV $ws 'B126' 'Proyectado plan'; SetV $ws 'B127' 'Meta 2027'
    SetF $ws 'C124' '=N(C8)'; SetF $ws 'C125' '=N(D44)'; SetF $ws 'C126' '=N(C44)'; SetF $ws 'C127' '=N(C9)'
    Cal $ws 'B124:C127'; Fmt $ws 'C124:C127' '0.00'

    SetV $ws 'B129' 'F · Auxiliares'; SetV $ws 'C129' 'Valor'; Hdr $ws 'B129:C129'
    SetV $ws 'B130' 'Unidad de la brecha'; SetF $ws 'C130' '=IF(C7="%","pp",IF(C7="","puntos",C7))'
    SetV $ws 'B131' 'Valor ideal usado'; SetF $ws 'C131' '=IF(C10="",IF(C6="Sube es mejor",100,0),C10)'
    SetV $ws 'B132' 'Posición del dolor principal'; SetF $ws 'C132' '=IF(MAX(C22:H22)=0,0,MATCH(MAX(C22:H22),C22:H22,0))'
    SetV $ws 'B133' 'Nombre del dolor principal'; SetF $ws 'C133' '=IF(C132=0,"",INDEX(C17:H17,C132))'
    Cal $ws 'B130:C133'

    # ---- example data
    if ($d) {
        SetV $ws 'C5' $d.name; SetV $ws 'C6' $d.sentido; SetV $ws 'C7' $d.unidad
        SetV $ws 'C8' $d.ejec; SetV $ws 'C9' $d.meta; SetV $ws 'C13' $d.total
        for ($i = 0; $i -lt $d.dol.Count; $i++) {
            $c = [char](67 + $i)
            SetV $ws "${c}17" $d.dol[$i][0]; SetV $ws "${c}18" $d.dol[$i][1]; SetV $ws "${c}23" $d.dol[$i][2]
        }
        for ($j = 0; $j -lt $d.ini.Count; $j++) {
            $r = 29 + $j; $it = $d.ini[$j]
            SetV $ws "B$r" $it.n
            foreach ($k in $it.e.Keys) { $c = [char](67 + [int]$k); SetV $ws "$c$r" $it.e[$k] }
            SetV $ws "I$r" $it.x; SetV $ws "J$r" $it.s; SetV $ws "K$r" $it.r
            SetV $ws "L$r" (OADate $it.f); if ($it.ev) { SetV $ws "M$r" $it.ev }
        }
    }

    # ---- chart
    $co = $ws.ChartObjects().Add($ws.Range('J40').Left + 5, $ws.Range('J40').Top, 440, 250)
    $ch = $co.Chart; $ch.ChartType = 51
    $ch.SetSourceData($ws.Range('B124:C127'), 2)
    $ch.HasTitle = $true; $ch.ChartTitle.Text = '¿Dónde quedaría el indicador en 2027?'; $ch.ChartTitle.Font.Size = 11
    $ch.HasLegend = $false
    $s = $ch.SeriesCollection(1); $s.HasDataLabels = $true; $s.DataLabels().NumberFormatLocal = FormatoLocal '0.00'
    $cols = @((RGB 166 166 166), (RGB 31 56 100), (RGB 155 194 230), (RGB 0 128 0))
    for ($p = 1; $p -le 4; $p++) { $s.Points($p).Format.Fill.ForeColor.RGB = $cols[$p - 1] }

    $ws.Protect($M, $true, $true, $true, $false, $true, $true, $true)
    NoGrid $x $ws 90
}

# ---------------------------------------------------------------- data
$ind1 = @{
    name = 'Churn mensual pospago'; sentido = 'Baja es mejor'; unidad = '%'; ejec = 2.8; meta = 2.0; total = 12000
    dol  = @(
        @('Precio / oferta de la competencia', 6000, 'Encuesta de salida Q2-2026 (n = 12.000 bajas)'),
        @('Fallas de red / cobertura', 3600, 'Cruce de bajas con mapa de caídas de red (últimos 90 días)'),
        @('Reclamos no resueltos', 3000, 'CRM: reclamo abierto más de 7 días antes de la baja'),
        @('Errores de facturación', 1800, 'Notas de crédito emitidas en los 3 meses previos a la baja'),
        @('No gestionable (mudanza, fallecimiento)', 1200, 'Motivo de baja declarado en tienda')
    )
    ini  = @(
        @{n = 'Nuevo portafolio de planes competitivos'; e = @{0 = 0.45 }; x = 0.40; s = 'Retrasada'; r = 'Marketing'; f = '2027-03-31'; ev = 'Piloto: 45 % de los clientes que se iban por precio aceptó el nuevo plan' },
        @{n = 'Ofertas de retención con modelo predictivo'; e = @{0 = 0.10; 2 = 0.15 }; x = 1.0; s = 'En plan'; r = 'Retención'; f = '2026-12-15'; ev = 'Piloto Q1-2026: retuvo 10 % (precio) y 15 % (reclamos) de clientes en riesgo' },
        @{n = 'Densificación de red en zonas críticas'; e = @{1 = 0.30 }; x = 0.60; s = 'En riesgo'; r = 'Red'; f = '2027-06-30'; ev = 'Simulación: 30 % de las quejas de red están en sitios planificados' },
        @{n = 'Resolución en el primer contacto'; e = @{2 = 0.30; 3 = 0.25 }; x = 0.90; s = 'En plan'; r = 'Experiencia cliente'; f = '2027-01-31'; ev = 'Piloto zona sur: −30 % reclamos reiterados y −25 % errores de factura' },
        @{n = 'Campaña de imagen de marca en redes'; e = @{0 = 0.01 }; x = 1.0; s = 'En plan'; r = 'Comunicaciones'; f = '2026-11-30'; ev = '' }
    )
}
$ind2 = @{
    name = 'Instalaciones en fecha (FTTH)'; sentido = 'Sube es mejor'; unidad = '%'; ejec = 82; meta = 92; total = 9000
    dol  = @(
        @('Falta de técnicos en la zona', 4050, 'Órdenes reprogramadas por falta de cuadrilla (sistema de despacho)'),
        @('Cliente ausente o dirección errónea', 2700, 'Visitas fallidas registradas por el técnico'),
        @('Falta de equipos (ONT / módem) en almacén', 1800, 'Órdenes en espera por stock (ERP de logística)'),
        @('Factibilidad mal validada en la venta', 1350, 'Órdenes rechazadas por falta de puerto o red')
    )
    ini  = @(
        @{n = 'Agendamiento con confirmación por WhatsApp'; e = @{1 = 0.70 }; x = 1.0; s = 'En plan'; r = 'Experiencia cliente'; f = '2026-12-31'; ev = 'Piloto zona norte: −70 % de visitas fallidas' },
        @{n = 'Contratación de técnicos tercerizados'; e = @{0 = 0.60 }; x = 0.90; s = 'En riesgo'; r = 'Operaciones'; f = '2027-02-28'; ev = 'Plan de capacidad: +60 % de cuadrillas en zonas con atraso' },
        @{n = 'Reposición automática de inventario'; e = @{2 = 0.80; 0 = 0.10 }; x = 1.0; s = 'En plan'; r = 'Logística'; f = '2027-01-15'; ev = 'Modelo de stock mínimo: cubre 80 % de los quiebres de stock' },
        @{n = 'Validación de factibilidad en el punto de venta'; e = @{3 = 0.70; 1 = 0.10 }; x = 1.0; s = 'En plan'; r = 'Ventas'; f = '2027-03-31'; ev = 'App de factibilidad: 70 % de los rechazos se detectan antes de vender' }
    )
}

# ---------------------------------------------------------------- run
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
try {
    $wb = $x.Workbooks.Add()
    try { $wb.Styles.Item('Normal').Font.Name = 'Arial'; $wb.Styles.Item('Normal').Font.Size = 10 } catch { }
    # Separadores y formato de fecha de este Excel
    $script:SEP_DEC = [string]$x.International(3)
    $script:SEP_MIL = [string]$x.International(4)
    $aux = $wb.Worksheets.Item(1).Range('A1')
    $aux.Formula = '=DATE(2027,3,31)'
    $script:FMT_FECHA = [string]$aux.NumberFormatLocal
    $aux.Clear() | Out-Null
    if ($script:FMT_FECHA -notmatch '[dDjJ]') { $script:FMT_FECHA = $null }
    $names = 'Guia', 'Resumen', 'IND01_Churn', 'IND02_Instalaciones', 'IND_PLANTILLA', 'Parametros'
    $prev = $wb.Worksheets.Item(1); $prev.Name = 'Guia'
    foreach ($n in $names[1..5]) { $s = $wb.Worksheets.Add($M, $prev); $s.Name = $n; $prev = $s }
    foreach ($s in @($wb.Worksheets)) { if ($names -notcontains $s.Name) { $s.Delete() } }
    $wb.SaveAs($out, 51)
    $x.Calculation = -4135  # manual during build

    # ------------------------------------------------ Parametros
    $ps = $wb.Worksheets.Item('Parametros')
    $ps.Cells.Font.Name = 'Arial'; $ps.Cells.Font.Size = 10
    Widths $ps @{A = 3; B = 40; C = 16; D = 90 }
    SetV $ps 'A1' 'Parámetros (ajustes generales)'; $ps.Range('A1').Font.Size = 16; $ps.Range('A1').Font.Bold = $true; $ps.Range('A1').Font.Color = $NAVY
    SetV $ps 'A2' 'Cambia estos valores solo si tu equipo lo acuerda. Afectan a todos los indicadores.'
    Band $ps 4 'D' 'Semáforo de suficiencia (de menor a mayor)'
    SetV $ps 'B5' 'Desde (suficiencia)'; SetV $ps 'C5' 'Color'; SetV $ps 'D5' 'Qué significa'; Hdr $ps 'B5:D5'
    SetV $ps 'B6' 0; SetV $ps 'C6' 'ROJO'; SetV $ps 'D6' 'Las iniciativas no alcanzan: faltan iniciativas o más efectividad.'
    SetV $ps 'B7' 0.7; SetV $ps 'C7' 'AMARILLO'; SetV $ps 'D7' 'Cerca: con ajustes puede alcanzar.'
    SetV $ps 'B8' 1; SetV $ps 'C8' 'VERDE'; SetV $ps 'D8' 'Alcanza: las iniciativas cubren toda la brecha.'
    Inp $ps 'B6:B8'; Cal $ps 'C6:D8'; Fmt $ps 'B6:B8' '0%'; CFSemaforo $ps 'C6:C8'
    Band $ps 10 'D' 'Alertas'
    SetV $ps 'B11' 'Parámetro'; SetV $ps 'C11' 'Valor'; SetV $ps 'D11' 'Qué significa'; Hdr $ps 'B11:D11'
    SetV $ps 'B12' 'Umbral para REEMPLAZAR'; SetV $ps 'C12' 0.05; SetV $ps 'D12' 'Si una iniciativa aporta menos de este % de la brecha (o su ejecución es 0 %), se marca REEMPLAZAR.'
    SetV $ps 'B13' 'Peso mínimo para SIN COBERTURA'; SetV $ps 'C13' 0.20; SetV $ps 'D13' 'Un dolor cuya atribución es mayor a este valor...'
    SetV $ps 'B14' 'Resolución mínima para SIN COBERTURA'; SetV $ps 'C14' 0.30; SetV $ps 'D14' '...y que queda resuelto menos que este valor (realista) se marca SIN COBERTURA.'
    SetV $ps 'B15' 'Tope de efectividad real'; SetV $ps 'C15' 0.99; SetV $ps 'D15' 'Ninguna iniciativa resuelve el 100 % de un dolor. También evita divisiones entre cero.'
    Cal $ps 'B12:B15'; Inp $ps 'C12:C15'; Cal $ps 'D12:D15'; Fmt $ps 'C12:C15' '0%'
    Band $ps 17 'D' 'Listas desplegables'
    SetV $ps 'B18' 'Sentido'; SetV $ps 'C18' 'Estado'; Hdr $ps 'B18:C18'
    SetV $ps 'B19' 'Sube es mejor'; SetV $ps 'B20' 'Baja es mejor'
    SetV $ps 'C19' 'En plan'; SetV $ps 'C20' 'En riesgo'; SetV $ps 'C21' 'Retrasada'
    Cal $ps 'B19:C21'
    $nm = @{Semaforo_Desde = '=Parametros!$B$6:$B$8'; Semaforo_Color = '=Parametros!$C$6:$C$8'; Umbral_Reemplazar = '=Parametros!$C$12'; Cobertura_Atrib_Min = '=Parametros!$C$13'; Cobertura_Resuelto_Min = '=Parametros!$C$14'; Tope_Efectividad = '=Parametros!$C$15' }
    foreach ($k in $nm.Keys) { $wb.Names.Add($k, $nm[$k]) | Out-Null }
    NoGrid $x $ps 100

    # ------------------------------------------------ indicator sheets
    New-HojaIndicador $x $wb.Worksheets.Item('IND01_Churn') $ind1 '⚠ DATOS DE EJEMPLO – reemplázalos por los tuyos.' $true
    New-HojaIndicador $x $wb.Worksheets.Item('IND02_Instalaciones') $ind2 '⚠ DATOS DE EJEMPLO – reemplázalos por los tuyos.' $true
    New-HojaIndicador $x $wb.Worksheets.Item('IND_PLANTILLA') $null 'Hoja en blanco: cópiala para cada indicador nuevo (clic derecho en la pestaña > Mover o copiar > Crear una copia).' $false

    # ------------------------------------------------ Resumen
    $rs = $wb.Worksheets.Item('Resumen')
    $rs.Cells.Font.Name = 'Arial'; $rs.Cells.Font.Size = 10; $rs.Cells.VerticalAlignment = -4108
    Widths $rs @{A = 8; B = 40; C = 30; D = 11; E = 11; F = 14; G = 11; H = 12; I = 12; J = 12; K = 12; L = 12; M = 15; N = 30; O = 13; P = 13; Q = 10 }
    SetV $rs 'A1' 'Resumen de suficiencia'; $rs.Range('A1').Font.Size = 16; $rs.Range('A1').Font.Bold = $true; $rs.Range('A1').Font.Color = $NAVY
    SetV $rs 'A2' 'Escribe en la columna B el nombre exacto de la hoja de cada indicador, en orden y sin dejar filas vacías. El resto se llena solo.'
    SetV $rs 'A3' 'Azul = lo escribes tú  ·  Gris = se calcula solo  ·  Plan = todas las iniciativas al 100 %  ·  Realista = con el % de ejecución esperado'
    $rs.Range('A3').Font.Italic = $true
    Band $rs 4 'Q' '1 · Comparativo por indicador'
    $h = @{A = 'N°'; B = 'Hoja del indicador'; C = 'Indicador'; D = 'Unidad'; E = 'Ejecutado hoy'; F = 'Meta 2027'; G = 'Brecha'; H = 'Proyectado · Plan'; I = 'Proyectado · Realista'; J = 'Brecha que faltaría (realista)'; K = 'Suficiencia · Plan'; L = 'Suficiencia · Realista'; M = 'Semáforo (realista)'; N = 'Dolor principal'; O = 'Iniciativas a REEMPLAZAR'; P = 'Iniciativas EN RIESGO'; Q = 'Referencia 100 %' }
    foreach ($k in $h.Keys) { SetV $rs "${k}5" $h[$k] }
    Hdr $rs 'A5:Q5'; $rs.Range('A5').RowHeight = 44
    for ($i = 0; $i -lt 10; $i++) { SetV $rs ("A" + (6 + $i)) ($i + 1) }
    $rs.Range('A6:A15').HorizontalAlignment = -4108; $rs.Range('A6:A15').Font.Bold = $true
    Inp $rs 'B6:B15'; Cal $rs 'C6:Q15'
    SetV $rs 'B6' 'IND01_Churn'; SetV $rs 'B7' 'IND02_Instalaciones'
    SetF $rs 'C6:C15' '=IF($B6="","",IFERROR(INDIRECT("''"&$B6&"''!C5"),"⚠ Hoja no encontrada"))'
    $map = @{D = 'C7'; E = 'C8'; F = 'C9'; G = 'C12'; H = 'C44'; I = 'D44'; J = 'D45'; K = 'C42'; L = 'D42'; M = 'D43'; N = 'C133' }
    foreach ($k in $map.Keys) { SetF $rs "${k}6:${k}15" ('=IF($B6="","",IFERROR(INDIRECT("''"&$B6&"''!' + $map[$k] + '"),""))') }
    SetF $rs 'O6:O15' '=IF($B6="","",IFERROR(COUNTIF(INDIRECT("''"&$B6&"''!F58:F65"),"REEMPLAZAR"),""))'
    SetF $rs 'P6:P15' '=IF($B6="","",IFERROR(COUNTIF(INDIRECT("''"&$B6&"''!F58:F65"),"EN RIESGO"),""))'
    SetF $rs 'Q6:Q15' '=IF($B6="","",1)'
    Fmt $rs 'E6:J15' '0.00'; Fmt $rs 'K6:L15' '0%'; Fmt $rs 'Q6:Q15' '0%'
    $rs.Range('D6:M15').HorizontalAlignment = -4108; $rs.Range('O6:Q15').HorizontalAlignment = -4108
    $rs.Range('K6:M15').Font.Bold = $true; $rs.Range('Q6:Q15').Font.Color = (RGB 128 128 128)
    CFSemaforo $rs 'M6:M15'; CFAlerts $rs 'C6:C15'

    Band $rs 17 'Q' '2 · ¿Cuánto cubren las iniciativas?  (la línea verde es 100 % = alcanza)'
    # Rangos dinámicos del gráfico: la fórmula se escribe en inglés y Excel la traduce a su idioma
    foreach ($nc in @(@('Res_Cat', 'C'), @('Res_Plan', 'K'), @('Res_Real', 'L'), @('Res_Ref', 'Q'))) {
        $nn = $wb.Names.Add($nc[0], '=Resumen!$' + $nc[1] + '$6')
        $nn.RefersToLocal = TraducirFormula $rs ('=OFFSET(Resumen!$' + $nc[1] + '$6,0,0,MAX(1,COUNTIF(Resumen!$B$6:$B$15,"?*")),1)')
    }
    $chTop = $rs.Range('B18').Top + 4
    $co = $rs.ChartObjects().Add($rs.Range('B18').Left, $chTop, 700, ($rs.Range('A36').Top - $chTop - 8))
    $ch = $co.Chart; $ch.ChartType = 51
    while ($ch.SeriesCollection().Count -gt 0) { $ch.SeriesCollection(1).Delete() }
    $bn = "'" + $wb.Name + "'!"
    $ser = @(@('Plan (todo al 100 %)', 'Res_Plan', (RGB 155 194 230)), @('Realista', 'Res_Real', (RGB 31 56 100)), @('100 % = alcanza', 'Res_Ref', (RGB 0 128 0)))
    foreach ($sd in $ser) {
        $s = $ch.SeriesCollection().NewSeries(); $s.Name = $sd[0]; $s.Values = '=' + $bn + $sd[1]; $s.XValues = '=' + $bn + 'Res_Cat'
        if ($sd[1] -eq 'Res_Ref') { $s.ChartType = 4; $s.Format.Line.ForeColor.RGB = $sd[2]; $s.Format.Line.Weight = 2; $s.Format.Line.DashStyle = 4 }
        else { $s.Format.Fill.ForeColor.RGB = $sd[2]; $s.HasDataLabels = $true; $s.DataLabels().NumberFormat = '0%' }
    }
    $ch.HasTitle = $true; $ch.ChartTitle.Text = 'Suficiencia por indicador'; $ch.ChartTitle.Font.Size = 11
    $ch.HasLegend = $true; $ch.Legend.Position = -4107
    $ch.Axes(2).TickLabels.NumberFormat = '0%'; $ch.Axes(2).MinimumScale = 0

    Band $rs 36 'Q' '3 · Ranking de iniciativas del indicador elegido'
    SetV $rs 'B37' 'Elige la hoja del indicador  →'; $rs.Range('B37').Font.Bold = $true; $rs.Range('B37').HorizontalAlignment = -4152
    Inp $rs 'C37'; SetV $rs 'C37' 'IND01_Churn'; DVList $rs 'C37' '=$B$6:$B$15'
    SetF $rs 'D37' '=IF(C37="","",IFERROR(INDIRECT("''"&C37&"''!C5"),"⚠ Hoja no encontrada"))'
    $rs.Range('D37').Font.Bold = $true; CFAlerts $rs 'D37'
    SetV $rs 'A38' 'Puesto'; SetV $rs 'B38' 'Iniciativa'; SetV $rs 'C38' 'Aporte (puntos)'; SetV $rs 'D38' 'Aporte (% brecha)'
    SetV $rs 'E38' 'Suficiencia si falla'; SetV $rs 'F38' 'Etiqueta'; SetV $rs 'G38' 'En palabras simples'
    Hdr $rs 'A38:G38'; $rs.Range('A38').RowHeight = 32
    for ($i = 0; $i -lt 8; $i++) {
        $r = 39 + $i; $src = 68 + $i
        SetV $rs "A$r" ($i + 1)
        foreach ($c in 'B', 'C', 'D', 'E', 'F', 'G') { SetF $rs "$c$r" ('=IF($C$37="","",IFERROR(INDIRECT("''"&$C$37&"''!' + $c + $src + '"),""))') }
    }
    $rs.Range('A39:A46').HorizontalAlignment = -4108; $rs.Range('A39:A46').Font.Bold = $true
    Cal $rs 'B39:F46'; Fmt $rs 'C39:C46' '0.00'; Fmt $rs 'D39:E46' '0%'; $rs.Range('C39:F46').HorizontalAlignment = -4108
    CFEtiqueta $rs 'F39:F46'

    Band $rs 48 'Q' '4 · Qué hacer con cada etiqueta'
    $et = @(@('APORTA', 'Va bien. Mantén el seguimiento normal.'), @('EN RIESGO', 'Aporta, pero está retrasada o en riesgo. Dale seguimiento cercano: si falla, la suficiencia cae.'), @('REEMPLAZAR', 'Casi no aporta (o su ejecución es 0 %). Cámbiala o rediséñala.'), @('⚠ SIN COBERTURA', 'Lo verás en la hoja del indicador: un dolor importante casi no se resuelve. Falta una iniciativa para él.'))
    for ($i = 0; $i -lt $et.Count; $i++) { $r = 49 + $i; SetV $rs "B$r" $et[$i][0]; SetV $rs "C$r" $et[$i][1]; $rs.Range("B$r").Font.Bold = $true; $rs.Range("B$r").HorizontalAlignment = -4108 }
    CFEtiqueta $rs 'B49:B51'; CFAlerts $rs 'B52'
    $rs.Protect($M, $true, $true, $true, $false, $true, $true, $true)
    NoGrid $x $rs 90

    # ------------------------------------------------ Guia
    $g = $wb.Worksheets.Item('Guia')
    $g.Cells.Font.Name = 'Arial'; $g.Cells.Font.Size = 10; $g.Cells.VerticalAlignment = -4108
    Widths $g @{A = 3; B = 34; C = 13; D = 13; E = 13; F = 13; G = 13; H = 13; I = 13; J = 13 }
    SetV $g 'A1' 'Plantilla de Suficiencia · Guía rápida'; $g.Range('A1').Font.Size = 18; $g.Range('A1').Font.Bold = $true; $g.Range('A1').Font.Color = $NAVY
    SetV $g 'A2' 'Lee esto una vez (3 minutos). Después solo sigues los 4 pasos de cada hoja de indicador.'
    $r = 4
    function Line($t, $bold) { SetV $g "B$script:r" $t; if ($bold) { $g.Range("B$script:r").Font.Bold = $true }; $script:r++ }
    function Pair($a, $b) { SetV $g "B$script:r" $a; SetV $g "C$script:r" $b; $g.Range("B$script:r").Font.Bold = $true; $script:r++ }
    function GBand($t) { $script:r++; Band $g $script:r 'J' $t; $script:r++ }
    $script:r = 3
    GBand '1 · ¿Qué es la suficiencia? La idea del pastel'
    Line '• El PASTEL es todo el problema de hoy: la distancia entre el valor actual y el valor ideal.'
    Line '• Cada DOLOR es una porción del pastel: una causa del problema, sacada del análisis de datos.'
    Line '• Cada INICIATIVA se come una parte de las porciones (dolores) que ataca.'
    Line '• Solo cuenta lo que la iniciativa de verdad ejecuta: si se retrasa, come menos.'
    Line '• La BRECHA es el pedazo que necesitas comer para llegar a la meta 2027.'
    Line '• SUFICIENCIA = lo que comen las iniciativas ÷ la brecha.   100 % o más = alcanza.' $true

    GBand '2 · Los 4 pasos (en cada hoja de indicador, de arriba hacia abajo)'
    Pair 'PASO 1 · ¿Cuánto me falta?' 'Escribe el ejecutado, la meta 2027 y el total de casos analizados.'
    Pair 'PASO 2 · ¿Qué lo causa?' 'Escribe los dolores y cuántos casos tiene cada uno. La plantilla calcula su peso en el pastel.'
    Pair 'PASO 3 · ¿Qué voy a hacer?' 'Escribe las iniciativas, qué % de cada dolor resuelven y qué % estará ejecutado a 2027.'
    Pair 'PASO 4 · ¿Me alcanza?' 'Solo miras: suficiencia, semáforo y qué iniciativa aporta más. Aquí no escribes nada.'

    GBand '3 · Colores'
    $c0 = $script:r
    Pair 'Azul' 'Lo escribes tú.'; Pair 'Gris' 'Se calcula solo. No lo toques.'; Pair '⚠ Rojo' 'Alerta: revisa ese dato.'; Pair '✔ Verde' 'Todo bien.'
    $g.Range("B$c0").Interior.Color = $INFILL; $g.Range("B$c0").Font.Color = $INFONT
    $g.Range("B$($c0+1)").Interior.Color = $GRAY
    $g.Range("B$($c0+2)").Interior.Color = $REDF; $g.Range("B$($c0+2)").Font.Color = $REDT
    $g.Range("B$($c0+3)").Interior.Color = $GRNF; $g.Range("B$($c0+3)").Font.Color = $GRNT

    GBand '4 · Ejemplo en 1 minuto  (cambia las celdas azules y mira qué pasa)'
    $e0 = $script:r
    SetV $g "B$e0" 'Problema total hoy (el pastel)'; SetV $g "C$e0" 10; SetV $g "D$e0" 'puntos'
    SetV $g "B$($e0+1)" 'Brecha a la meta'; SetV $g "C$($e0+1)" 4; SetV $g "D$($e0+1)" 'puntos'
    SetV $g "B$($e0+2)" 'Iniciativa X: se ejecuta al'; SetV $g "C$($e0+2)" 1
    SetV $g "B$($e0+3)" 'Iniciativa Y: se ejecuta al'; SetV $g "C$($e0+3)" 0.5
    Inp $g "C${e0}:C$($e0+3)"; Fmt $g "C$($e0+2):C$($e0+3)" '0%'; $g.Range("B${e0}:B$($e0+3)").Font.Bold = $true
    $th = $e0 + 5
    $hd = @('Dolor', 'Atribución (peso)', 'X resuelve', 'Y resuelve', 'X real (× ejecución)', 'Y real (× ejecución)', 'Lo que queda del dolor', '% resuelto', 'Puntos recuperados')
    for ($i = 0; $i -lt 9; $i++) { SetV $g ("" + [char](66 + $i) + $th) $hd[$i] }
    Hdr $g "B${th}:J${th}"; $g.Range("A$th").RowHeight = 32
    $a = $th + 1; $b = $th + 2
    SetV $g "B$a" 'Dolor A'; SetV $g "C$a" 0.6; SetV $g "D$a" 0.5; SetV $g "E$a" 0.2
    SetV $g "B$b" 'Dolor B'; SetV $g "C$b" 0.4; SetV $g "D$b" 0; SetV $g "E$b" 0.5
    Inp $g "B${a}:E${b}"; Cal $g "F${a}:J${b}"
    SetF $g "F${a}:F${b}" ('=D' + $a + '*$C$' + ($e0 + 2))
    SetF $g "G${a}:G${b}" ('=E' + $a + '*$C$' + ($e0 + 3))
    SetF $g "H${a}:H${b}" ('=(1-F' + $a + ')*(1-G' + $a + ')')
    SetF $g "I${a}:I${b}" ('=1-H' + $a)
    SetF $g "J${a}:J${b}" ('=$C$' + $e0 + '*C' + $a + '*I' + $a)
    Fmt $g "C${a}:I${b}" '0%'; Fmt $g "J${a}:J${b}" '0.00'; $g.Range("C${a}:J${b}").HorizontalAlignment = -4108
    $t1 = $b + 1; $t2 = $b + 2; $t3 = $b + 3
    SetV $g "B$t1" 'Total de puntos recuperados'; SetF $g "J$t1" ('=SUM(J' + $a + ':J' + $b + ')')
    SetV $g "B$t2" 'Suficiencia = recuperado ÷ brecha'; SetF $g "J$t2" ('=IF(C' + ($e0 + 1) + '=0,0,J' + $t1 + '/C' + ($e0 + 1) + ')')
    SetV $g "B$t3" 'Semáforo'; SetF $g "J$t3" ('=LOOKUP(J' + $t2 + ',Semaforo_Desde,Semaforo_Color)')
    $g.Range("B${t1}:B${t3}").Font.Bold = $true; Cal $g "J${t1}:J${t3}"; Fmt $g "J$t1" '0.00'; Fmt $g "J$t2" '0%'
    $g.Range("J${t1}:J${t3}").Font.Bold = $true; $g.Range("J${t1}:J${t3}").HorizontalAlignment = -4108
    CFSemaforo $g "J$t3"
    $script:r = $t3 + 2
    SetF $g "B$script:r" ('="Resultado: las iniciativas recuperan "&FIXED(J' + $t1 + ',1)&" puntos de una brecha de "&FIXED(C' + ($e0 + 1) + ',1)&"  →  suficiencia "&FIXED(J' + $t2 + '*100,0)&" %."'); $g.Range("B$script:r").Font.Bold = $true; $script:r++
    Line '¿Por qué se multiplica "lo que queda"? Porque la segunda iniciativa solo puede resolver lo que la primera dejó.'
    SetF $g "B$script:r" ('="En el dolor A: X deja el "&FIXED((1-F' + $a + ')*100,0)&" %; luego Y deja el "&FIXED((1-G' + $a + ')*100,0)&" % de eso  →  queda "&FIXED(H' + $a + '*100,0)&" %  →  se resolvió el "&FIXED(I' + $a + '*100,0)&" %. Así un dolor nunca pasa de 100 %."'); $script:r++
    Line 'Prueba: cambia la ejecución de Y a 0 % y mira cómo baja la suficiencia.'

    GBand '5 · ¿Por qué la suficiencia puede pasar de 100 %?'
    Line 'Porque los dolores explican TODO el problema (el pastel), no solo la brecha. Si resuelves más de lo que pide la meta, sobra.'
    Line 'Ejemplo: el churn hoy es 2,8 % y la meta 2,0 %. Los dolores explican los 2,8 puntos; basta recuperar 0,8 para cumplir.'

    GBand '6 · Glosario'
    $glos = @(
        @('Ejecutado', 'Valor real del indicador hoy.'),
        @('Meta 2027', 'Valor que debes alcanzar en 2027.'),
        @('Valor ideal', 'El mejor valor posible (normalmente 100 % o 0 %).'),
        @('Problema total (pastel)', 'Distancia entre el ejecutado y el valor ideal.'),
        @('Brecha', 'Distancia entre el ejecutado y la meta. Es lo que debes recuperar.'),
        @('Dolor', 'Una causa del problema, sacada del análisis de datos del indicador.'),
        @('Prevalencia', '% de casos que tienen ese dolor. La suma puede pasar de 100 % (un caso puede tener varios dolores).'),
        @('Atribución', 'Parte del pastel que le toca a cada dolor. Siempre suma 100 %.'),
        @('Efectividad', '% del dolor que la iniciativa resuelve si se hace completa. Lo defines tú, con evidencia.'),
        @('Ejecución', '% de la iniciativa que de verdad estará lista a 2027 (retrasos o alcance parcial).'),
        @('Suficiencia', 'Puntos recuperados ÷ brecha. 100 % o más = las iniciativas alcanzan.'),
        @('Aporte', 'Cuánto baja la suficiencia si esa iniciativa falla.')
    )
    foreach ($p in $glos) { Pair $p[0] $p[1] }

    GBand '7 · Qué significan las etiquetas'
    $l0 = $script:r
    Pair 'APORTA' 'Va bien. Seguimiento normal.'
    Pair 'EN RIESGO' 'Aporta, pero está retrasada o en riesgo. Si falla, la suficiencia cae: dale seguimiento cercano.'
    Pair 'REEMPLAZAR' 'Casi no aporta (menos del 5 % de la brecha) o su ejecución es 0 %. Cámbiala o rediséñala.'
    Pair '⚠ SIN COBERTURA' 'Un dolor importante casi no se resuelve. Falta una iniciativa para él.'
    Pair '⚠ Sin respaldo' 'Falta la evidencia que justifica el %. Toda cifra debe tener su fuente de datos.'
    $g.Range("B${l0}:B$($l0+4)").HorizontalAlignment = -4108
    CFEtiqueta $g "B${l0}:B$($l0+2)"; CFAlerts $g "B$($l0+3):B$($l0+4)"

    GBand '8 · Cómo agregar un indicador (3 pasos)'
    Pair '1 · Copia la plantilla' 'Clic derecho en la pestaña IND_PLANTILLA > Mover o copiar > Crear una copia. Renómbrala sin espacios (ej. IND03_Disponibilidad).'
    Pair '2 · Llena las celdas azules' 'Pasos 1, 2 y 3 de la hoja nueva. El Paso 4 se calcula solo.'
    Pair '3 · Agrégala al Resumen' 'En la hoja Resumen, escribe el nombre de la hoja en la siguiente fila libre de la columna B.'

    GBand '9 · Reglas para no romper la plantilla'
    Line '• No insertes ni borres filas o columnas en las hojas de indicador (el Resumen las lee por posición).'
    Line '• En el Paso 1 escribe solo números (2,8 y no 2,8 %). En los Pasos 2 y 3 los porcentajes sí llevan % (40 %).'
    Line '• Si cambias una atribución a mano, cámbialas todas para que sumen 100 %.'
    Line '• Máximo 6 dolores y 8 iniciativas por indicador (cabe en una pantalla).'
    Line '• Las hojas están protegidas sin contraseña para cuidar las fórmulas: Revisar > Desproteger hoja.'

    $g.Protect($M, $true, $true, $true, $false, $true, $true, $true)
    NoGrid $x $g 100

    # ------------------------------------------------ tabs, finish
    $wb.Worksheets.Item('Guia').Tab.Color = (RGB 0 128 0)
    $wb.Worksheets.Item('Resumen').Tab.Color = $NAVY
    $wb.Worksheets.Item('IND01_Churn').Tab.Color = (RGB 91 155 213)
    $wb.Worksheets.Item('IND02_Instalaciones').Tab.Color = (RGB 91 155 213)
    $wb.Worksheets.Item('IND_PLANTILLA').Tab.Color = (RGB 166 166 166)
    $wb.Worksheets.Item('Parametros').Tab.Color = (RGB 166 166 166)
    $x.Calculation = -4105
    $x.CalculateFull()
    foreach ($s in @($wb.Worksheets)) { $s.Activate(); $s.Range('A1').Select() | Out-Null }
    $wb.Worksheets.Item('Guia').Activate()
    $x.ScreenUpdating = $true
    $wb.Save()
    'GUARDADO: ' + $out
}
finally {
    if ($wb) { $wb.Close($false) }
    $x.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
}
