# Verifica Plantilla_Suficiencia.xlsx con Excel (requiere Microsoft Excel en Windows).
# Recalcula todo, busca errores de fórmula, compara los ejemplos con los valores esperados
# (los mismos de tests/paridad_excel.test.js) y prueba copiar la hoja IND_PLANTILLA.
# El archivo se abre en solo lectura: nunca se modifica.
# Uso:  powershell -ExecutionPolicy Bypass -File herramientas\verificar_plantilla.ps1 [-Archivo ruta\archivo.xlsx]
param([string]$Archivo = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Plantilla_Suficiencia.xlsx'))
$ErrorActionPreference = 'Stop'
$ruta = [System.IO.Path]::GetFullPath($Archivo)
$script:fallas = 0

function Revisar($ok, $mensaje) {
    if ($ok) { "OK     $mensaje" } else { "FALLA  $mensaje"; $script:fallas++ }
}
function Cerca($a, $b) {
    if (-not ($a -is [double] -or $a -is [int])) { return $false }
    return [math]::Abs([double]$a - [double]$b) -le 1e-9 * [math]::Max(1, [math]::Abs([double]$b))
}
# Escribe números con Formula en formato invariante (evita problemas de configuración regional)
function Numero($rango, $valor) { $rango.Formula = ([double]$valor).ToString([System.Globalization.CultureInfo]::InvariantCulture) }

$esperado = [ordered]@{
    'IND01_Churn' = @{ C11 = 2.8; C12 = 0.79999999999999982; C41 = 1.0418692307692305; D41 = 0.68344769230769231; C42 = 1.3023365384615384; D42 = 0.85430961538461558; C43 = 'VERDE'; D43 = 'AMARILLO'; B68 = 'Resolución en el primer contacto' }
    'IND02_Instalaciones' = @{ C11 = 18; C12 = 10; C41 = 12.632727272727275; D41 = 12.235090909090912; C42 = 1.2632727272727275; D42 = 1.2235090909090913; C43 = 'VERDE'; D43 = 'VERDE'; B68 = 'Contratación de técnicos tercerizados' }
}

if (-not (Test-Path $ruta)) { "No existe el archivo $ruta"; exit 1 }
"Verificando $ruta`n"
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($ruta, 0, $true)
    $x.CalculateFull()

    foreach ($s in @($wb.Worksheets)) {
        $n = 0
        try { $n = $s.UsedRange.SpecialCells(-4123, 16).Count } catch { $n = 0 }
        Revisar ($n -eq 0) "$($s.Name): sin errores de fórmula"
    }

    foreach ($hoja in $esperado.Keys) {
        $ws = $wb.Worksheets.Item($hoja)
        $e = $esperado[$hoja]
        foreach ($k in 'C11', 'C12', 'C41', 'D41', 'C42', 'D42') { Revisar (Cerca $ws.Range($k).Value2 $e[$k]) "$hoja $k = $($ws.Range($k).Value2)" }
        foreach ($k in 'C43', 'D43', 'B68') { Revisar ($ws.Range($k).Value2 -eq $e[$k]) "$hoja $k = $($ws.Range($k).Value2)" }
    }

    # Copiar la hoja plantilla, llenarla y verla en el Resumen (sin guardar)
    $p = $wb.Worksheets.Item('IND_PLANTILLA')
    $p.Copy([Type]::Missing, $p)
    $c = $wb.Worksheets.Item($p.Index + 1)
    $c.Name = 'IND_TEST'
    $c.Range('C5').Value2 = 'Prueba'
    $c.Range('C6').Value2 = 'Sube es mejor'
    $c.Range('C7').Value2 = '%'
    Numero $c.Range('C8') 90
    Numero $c.Range('C9') 95
    Numero $c.Range('C13') 100
    $c.Range('C17').Value2 = 'Dolor X'
    Numero $c.Range('C18') 100
    $c.Range('B29').Value2 = 'Iniciativa X'
    Numero $c.Range('C29') 0.5
    Numero $c.Range('I29') 1
    $c.Range('M29').Value2 = 'dato'
    $rs = $wb.Worksheets.Item('Resumen')
    $rs.Unprotect()
    $rs.Range('B8').Value2 = 'IND_TEST'
    $x.CalculateFull()
    Revisar (Cerca $c.Range('D42').Value2 1) "copia de IND_PLANTILLA: suficiencia = $($c.Range('D42').Value2) (esperado 1 = 100 %)"
    Revisar ($rs.Range('C8').Value2 -eq 'Prueba') "el Resumen lee la hoja copiada ($($rs.Range('C8').Value2))"
}
finally {
    if ($wb) { $wb.Close($false) }
    $x.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
}
if ($script:fallas) { "`n$($script:fallas) verificaciones fallaron."; exit 1 }
"`nTodo correcto."
