param(
    [Parameter(HelpMessage="URL a abrir en Google Chrome")]
    [string]$Url = "https://www.example.com",

    [Parameter(HelpMessage="Intervalo en segundos para refrescar la página")]
    [int]$IntervalSeconds = 60
)

$ErrorActionPreference = 'Stop'

function Get-ChromePath {
    $candidates = @(
        (Join-Path $env:ProgramFiles 'Google/Chrome/Application/chrome.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Google/Chrome/Application/chrome.exe'),
        (Join-Path $env:LocalAppData 'Google/Chrome/Application/chrome.exe')
    )
    foreach ($p in $candidates) {
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    return 'chrome.exe'
}

function Start-ChromeWindow {
    param(
        [Parameter(Mandatory=$true)][string]$TargetUrl
    )
    $chromePath = Get-ChromePath

    $args = @('--new-window', $TargetUrl)

    try {
        return Start-Process -FilePath $chromePath -ArgumentList $args -PassThru -WindowStyle Normal
    } catch {
        throw "No se pudo iniciar Google Chrome. Asegúrate de que esté instalado. Detalle: $($_.Exception.Message)"
    }
}

function Refresh-ChromeWindow {
    param(
        [int]$ProcessId
    )
    try {
        $wshell = New-Object -ComObject WScript.Shell
    } catch {
        throw "No se pudo crear el objeto COM WScript.Shell. Este script requiere Windows Script Host habilitado."
    }

    $activated = $false
    if ($ProcessId -gt 0) {
        try { $activated = $wshell.AppActivate($ProcessId) } catch { $activated = $false }
    }
    if (-not $activated) {
        try { $activated = $wshell.AppActivate('Google Chrome') } catch { $activated = $false }
    }

    if ($activated) {
        Start-Sleep -Milliseconds 150
        $wshell.SendKeys('{F5}') | Out-Null
        return $true
    }
    return $false
}

Write-Host "Abriendo Google Chrome con la última sesión de usuario y navegando a: $Url" -ForegroundColor Cyan
$chromeProc = Start-ChromeWindow -TargetUrl $Url

Start-Sleep -Seconds 3

# Bucle de auto-refresco
while ($true) {
    Start-Sleep -Seconds $IntervalSeconds

    if ($null -ne $chromeProc) {
        try {
            if ($chromeProc.HasExited) { $chromeProc = $null }
        } catch { $chromeProc = $null }
    }

    if ($null -eq $chromeProc) {
        $chromeProc = Get-Process -Name chrome -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle } |
            Sort-Object StartTime -Descending |
            Select-Object -First 1
        if ($null -eq $chromeProc) {
            Write-Host 'No hay ventana de Chrome activa. Finalizando script.' -ForegroundColor Yellow
            break
        }
    }

    $ok = Refresh-ChromeWindow -ProcessId $chromeProc.Id
    if (-not $ok) {
        Write-Host 'No se pudo activar la ventana de Chrome para refrescar. Reintentando en el próximo ciclo…' -ForegroundColor Yellow
    }
}
