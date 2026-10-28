# Prueba el tablero (index.html) en un navegador sin ventana (Microsoft Edge o Google Chrome).
#  1. Interacción: deslizadores, celdas, deshacer, cambiar de indicador, crear uno nuevo, guardar.
#  2. Importar Excel: carga Plantilla_Suficiencia.xlsx y compara con los ejemplos (necesita internet,
#     porque el lector de Excel se descarga de cdnjs).
# Uso:  powershell -ExecutionPolicy Bypass -File herramientas\probar_tablero.ps1 [-SinImportar]
param([switch]$SinImportar)
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
$html = [System.IO.File]::ReadAllText((Join-Path $raiz 'index.html'), [System.Text.Encoding]::UTF8)
$html = $html.Replace('<head>', '<head><base href="' + $urlRaiz + '">')

function Correr($nombre, $script, $extra) {
    $pagina = Join-Path $temp ($nombre + '.html')
    $conPrueba = $html.Replace('</body>', '<script src="' + $urlRaiz + 'tests/navegador/' + $script + '"></script></body>')
    [System.IO.File]::WriteAllText($pagina, $conPrueba, (New-Object System.Text.UTF8Encoding($false)))
    $salida = Join-Path $temp ($nombre + '-dom.txt')
    $perfil = Join-Path $temp ('perfil-' + $nombre)
    $argumentos = @('--headless=new', '--disable-gpu', "--user-data-dir=$perfil", '--virtual-time-budget=30000') + $extra + @('--dump-dom', ([System.Uri]$pagina).AbsoluteUri)
    Start-Process -FilePath $nav -ArgumentList $argumentos -RedirectStandardOutput $salida -RedirectStandardError (Join-Path $temp ($nombre + '-err.txt')) -Wait -NoNewWindow
    $dom = [System.IO.File]::ReadAllText($salida)
    if ($dom -match '(?s)<pre id="resultado-prueba">(.*?)</pre>') { return [System.Net.WebUtility]::HtmlDecode($Matches[1]) }
    return 'FALLA  la prueba no terminó (revisa que la página cargue sin errores)'
}

$resultados = @()
"== Interacción"
$r = Correr 'interaccion' 'interaccion.js' @()
$r
$resultados += $r
if (-not $SinImportar) {
    "`n== Importar Excel"
    $r = Correr 'importar' 'importar_excel.js' @('--allow-file-access-from-files')
    $r
    $resultados += $r
}
Remove-Item -Recurse -Force $temp -ErrorAction SilentlyContinue
$texto = $resultados -join "`n"
$fallas = ([regex]::Matches($texto, '(?m)^(FALLA|EXCEPCION)')).Count
$errores = $texto -notmatch 'errores JS: \[\]'
if ($fallas -or $errores) { "`nHay $fallas pruebas con falla o errores de JavaScript."; exit 1 }
"`nTodo correcto: " + ([regex]::Matches($texto, '(?m)^OK')).Count + ' comprobaciones.'
