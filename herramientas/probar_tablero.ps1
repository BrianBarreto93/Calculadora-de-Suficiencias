# Prueba el tablero (suficiencia_cx.html) en un navegador sin ventana (Microsoft Edge o Google Chrome).
#  1. Interacción: deslizadores, celdas, deshacer, nuevo cálculo desde cero, 15 dolores, guardar.
#  2. Importar Excel: carga Plantilla_Suficiencia.xlsx y compara con los ejemplos.
#  3. Exportar Excel: exporta, relee el archivo, lo edita como lo haría alguien en Excel y lo vuelve a importar.
#     Si Excel está instalado, además abre el archivo exportado con Excel para confirmar que no pide reparación.
#  4. Animaciones: los pasteles se despliegan y terminan en el valor correcto.
# Las pruebas 1 a 3 corren con "reducir movimiento" para leer los valores sin esperar las animaciones.
# Uso:  powershell -ExecutionPolicy Bypass -File herramientas\probar_tablero.ps1 [-SinImportar] [-SinExcel]
param([switch]$SinImportar, [switch]$SinExcel)
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$navegadores = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
$nav = $navegadores | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if (-not $nav) { 'No encontré Microsoft Edge ni Google Chrome.'; exit 1 }

$temp = Join-Path ([System.IO.Path]::GetTempPath()) ('suficiencia-prueba-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force $temp | Out-Null
$urlRaiz = ([System.Uri](Join-Path $raiz '.')).AbsoluteUri
if (-not $urlRaiz.EndsWith('/')) { $urlRaiz += '/' }
$html = [System.IO.File]::ReadAllText((Join-Path $raiz 'suficiencia_cx.html'), [System.Text.Encoding]::UTF8)
$html = $html.Replace('<head>', '<head><base href="' + $urlRaiz + '">')
$sinMovimiento = '--force-prefers-reduced-motion'

function Correr($nombre, $script, $extra) {
    $pagina = Join-Path $temp ($nombre + '.html')
    $conPrueba = $html.Replace('</body>', '<script src="' + $urlRaiz + 'tests/navegador/' + $script + '"></script></body>')
    [System.IO.File]::WriteAllText($pagina, $conPrueba, (New-Object System.Text.UTF8Encoding($false)))
    $salida = Join-Path $temp ($nombre + '-dom.txt')
    $perfil = Join-Path $temp ('perfil-' + $nombre)
    $argumentos = @('--headless=new', '--disable-gpu', "--user-data-dir=$perfil", '--virtual-time-budget=30000') + $extra + @('--dump-dom', ([System.Uri]$pagina).AbsoluteUri)
    Start-Process -FilePath $nav -ArgumentList $argumentos -RedirectStandardOutput $salida -RedirectStandardError (Join-Path $temp ($nombre + '-err.txt')) -Wait -NoNewWindow
    $dom = [System.IO.File]::ReadAllText($salida)
    $script:ultimoDom = $dom
    if ($dom -match '(?s)<pre id="resultado-prueba">(.*?)</pre>') { return [System.Net.WebUtility]::HtmlDecode($Matches[1]) }
    return 'FALLA  la prueba no terminó (revisa que la página cargue sin errores)'
}

# Abre con Excel (solo lectura) el archivo que exportó el tablero y revisa sus hojas y algunos valores.
function RevisarConExcel($ruta) {
    $lineas = @()
    try { $x = New-Object -ComObject Excel.Application } catch { return @('Excel no está instalado: se omite la revisión con Excel.') }
    $x.Visible = $false
    $x.DisplayAlerts = $false
    try {
        try { $wb = $x.Workbooks.Open($ruta, 0, $true) }
        catch { return @('FALLA  Excel no pudo abrir el archivo exportado (estaría dañado)') }
        $nombres = @($wb.Worksheets | ForEach-Object { $_.Name })
        $lineas += 'OK    Excel abre el archivo exportado: ' + ($nombres -join ' | ')
        $res = $wb.Worksheets.Item('Resumen')
        $lineas += $(if ($res.Range('A1').Value2 -eq 'SUFICIENCIA 2027 · RESUMEN') { 'OK    ' } else { 'FALLA ' }) + 'Resumen A1 = ' + $res.Range('A1').Value2
        $churn = $wb.Worksheets.Item(2)
        $celda = $churn.UsedRange.Find('Suficiencia esperada')
        $texto = if ($celda) { $churn.Cells($celda.Row, 3).Text } else { '' }
        $lineas += $(if ($texto -match '^85\s?%$') { 'OK    ' } else { 'FALLA ' }) + 'Excel muestra la suficiencia del Churn → ' + $texto
        $wb.Close($false)
    } finally {
        $x.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
    }
    return $lineas
}

$resultados = @()
"== Interacción"
$r = Correr 'interaccion' 'interaccion.js' @($sinMovimiento)
$r
$resultados += $r
if (-not $SinImportar) {
    "`n== Importar Excel (plantilla)"
    $r = Correr 'importar' 'importar_excel.js' @($sinMovimiento, '--allow-file-access-from-files')
    $r
    $resultados += $r
}
"`n== Exportar Excel"
$r = Correr 'exportar' 'exportar_excel.js' @($sinMovimiento)
$r
$resultados += $r
if (-not $SinExcel -and $script:ultimoDom -match '(?s)<pre id="xlsx-base64">(.*?)</pre>') {
    $archivo = Join-Path $temp 'exportado.xlsx'
    [System.IO.File]::WriteAllBytes($archivo, [System.Convert]::FromBase64String($Matches[1].Trim()))
    $r = (RevisarConExcel $archivo) -join "`n"
    $r
    $resultados += $r
}
"`n== Animaciones"
$r = Correr 'animaciones' 'animaciones.js' @()
$r
$resultados += $r
Remove-Item -Recurse -Force $temp -ErrorAction SilentlyContinue
$texto = $resultados -join "`n"
$fallas = ([regex]::Matches($texto, '(?m)^(FALLA|EXCEPCION)')).Count
$conErrores = ([regex]::Matches($texto, 'errores JS: \[\]')).Count -lt ([regex]::Matches($texto, 'errores JS:')).Count
if ($fallas -or $conErrores) { "`nHay $fallas pruebas con falla o errores de JavaScript."; exit 1 }
"`nTodo correcto: " + ([regex]::Matches($texto, '(?m)^OK')).Count + ' comprobaciones.'
