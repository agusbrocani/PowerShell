# ==========================================
# CONFIGURACION DE LOG
# ==========================================
$RutaLog = "C:/Users/PC/Desktop/"
$NombreBase = "NombreRealDelLog"
$FechaHoraActual = Get-Date -Format "dd-MM-yyyy_HH-mm"
$NombreLog = "${NombreBase}_${FechaHoraActual}hs.log"
$RutaCompleta = Join-Path $RutaLog $NombreLog

# ==========================================
# FUNCION PARA ESCRIBIR EN LOG (se debe aclarar al VERBO al llamarla)
# ==========================================
function Write-Log {
    param (
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Nivel = "INFO",
        [Parameter(Mandatory)][string]$Mensaje
    )

    $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss.fff"
    $linea = "[$timestamp] [$Nivel] $Mensaje"

    try {
        $global:stream.WriteLine($linea)
    } catch {
        Write-Host "Error al escribir log: $_" -ForegroundColor Red
    }
}

# ==========================================
# INICIO DE LOGGING Y CONEXION
# ==========================================
try {
    $inicio = Get-Date

    # Crear directorio si no existe
    if (-not (Test-Path -Path $RutaLog)) {
        New-Item -Path $RutaLog -ItemType Directory -Force | Out-Null
    }

    $global:stream = [System.IO.StreamWriter]::new($RutaCompleta, $true, [System.Text.Encoding]::UTF8)

    Write-Log -Nivel "INFO" -Mensaje "Inicio del script"
    Write-Host "Conectando a SharePoint Online..." -ForegroundColor Magenta
    Write-Log -Nivel "INFO" -Mensaje "Conectando a SharePoint Online..."

    $SiteURL = "https://circo.sharepoint.com/sites/SPXTest/"
    $WarningPreference = "SilentlyContinue"
    Connect-PnPOnline -Url $SiteURL -UseWebLogin
    Write-Log -Nivel "INFO" -Mensaje "Conectado exitosamente a $SiteURL"

    # ==========================================
    # ACCION PRINCIPAL (personalizable)
    # ==========================================
    Write-Host "Ejecutando acción principal..." -ForegroundColor Cyan
    Write-Log -Nivel "INFO" -Mensaje "Inicio de acción principal"

    # LÓGICA DE LA ACCIÓN PRINCIPAL
    for ($i = 1; $i -le 5; $i++) {
        $mensaje = "Registro de prueba $i"
        Write-Host "$mensaje" -ForegroundColor DarkGray
        Write-Log -Nivel "INFO" -Mensaje $mensaje
    }

    # ==========================================
    # FIN DE ACCIÓN PRINCIPAL
    # ==========================================
    Write-Host "Script ejecutado satisfactoriamente." -ForegroundColor Green
    Write-Log -Nivel "INFO" -Mensaje "Script ejecutado satisfactoriamente"

    $fin = Get-Date
    $duracion = $fin - $inicio
    $duracionFormateada = "{0:00}:{1:00}:{2:00}.{3:000}" -f $duracion.Hours, $duracion.Minutes, $duracion.Seconds, $duracion.Milliseconds

    Write-Log -Nivel "INFO" -Mensaje "Duración total del script [hh:mm:ss]: $duracionFormateada"
    Write-Host "Duración total del script [hh:mm:ss]: $duracionFormateada" -ForegroundColor Green
}
catch {
    Write-Host "¡Error! $($_.Exception.Message)" -ForegroundColor Red
    Write-Log -Nivel "ERROR" -Mensaje $_.Exception.Message
}
finally {
    Disconnect-PnPOnline
    Write-Host "Conexión cerrada." -ForegroundColor Magenta
    Write-Log -Nivel "INFO" -Mensaje "Conexión cerrada"

    if ($global:stream) {
        $global:stream.Close()
    }
}

# Uso apropiado de colores
    # Magenta   -> recursos
    # Cyan      -> acciones
    # Green     -> acciones completadas con éxito
    # Red       -> errores
    # Yellow    -> warnings
    # DarkGrey  -> uso común

# ESCRIBIR en archivo log:
    # Write-Log -Nivel "ERROR" -Mensaje $mensajeAEscribir
    # -Nivel    [OPCIONAL]. Valores posibles: "INFO", "WARNING", "ERROR". Por defecto es "INFO". 
    # -Mensaje  [OBLIGATORIO]
