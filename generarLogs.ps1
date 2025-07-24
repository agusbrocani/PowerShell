# CONFIGURACION
$RutaLog = "C:/Users/PC/Desktop/"
$nombreBase = "PRODUCTOS"

# Obtener fecha y hora actual con formato personalizado
$fechaHoraActual = Get-Date -Format "dd-MM-yyyy_HH-mm"
$NombreLog = "$nombreBase" + "_" + "$fechaHoraActual" + "hs" + ".log"

$RutaCompleta = Join-Path $RutaLog $NombreLog

# Crear directorio si no existe
if (-not (Test-Path -Path $RutaLog)) {
    New-Item -Path $RutaLog -ItemType Directory -Force | Out-Null
}

try {
    # Abrir el archivo una sola vez para escritura continua
    $stream = [System.IO.StreamWriter]::new($RutaCompleta, $true, [System.Text.Encoding]::UTF8)

    # Simulacion: escribir 10000 registros
    for ($i = 1; $i -le 10000; $i++) {
        $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss.fff"
        $linea = "[$timestamp] Registro número $i`n`t"
        $stream.WriteLine($linea)
    }

    # Cerrar el stream una sola vez
    $stream.Close()
}
catch {
    Write-Host "❌ Error escribiendo el log masivo: $_" -ForegroundColor Red
}
