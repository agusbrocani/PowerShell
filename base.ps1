param (
    [Parameter(Mandatory)][string]$SiteURL#,
    # [Parameter(Mandatory)][string]$ClientID,
    # [Parameter(Mandatory)][string]$ClientSecret
)

# ==========================================
# CONFIGURACION DE LOG
# ==========================================
$RutaLog = "C:/logs/"
$NombreBase = "CampoActivo"
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
    }
    catch {
        Write-Host "Error al escribir log: $_" -f Red
    }
}

# ==========================================
# INICIO DE LOGGING Y CONEXION
# ==========================================

$DisplayName = "Activo"
$InternalName = "Activo"
$SiteName = "OBIRAS"

Try {
    $inicio = Get-Date

    if (-not (Test-Path -Path $RutaLog)) {
        New-Item -Path $RutaLog -ItemType Directory -Force | Out-Null
    }

    $global:stream = [System.IO.StreamWriter]::new($RutaCompleta, $true, [System.Text.Encoding]::UTF8)

    Write-Log -Mensaje "Inicio del script"
    Write-Log -Mensaje "Conectando a SharePoint Online..."
    Write-Host "Conectando a SharePoint Online..." -f Magenta
    $WarningPreference = "SilentlyContinue"
    Connect-PnPOnline -Url $SiteURL -UseWebLogin

    Write-Log -Mensaje "Obteniendo todas las listas del sitio '$SiteName'..."
    Write-Host "Obteniendo todas las listas del sitio '$SiteName'...`n" -f Cyan
    $lists = Get-PnPList

    $obirasListNames = @()
    $obirasCounter = 0
    foreach ($list in $lists) {
        if ($list.Title.StartsWith("OBIRAS")) {
            $obirasListNames += $list.Title
            $obirasCounter++
            Write-Log -Mensaje "Lista encontrada: $($list.Title)"
            Write-Host "Lista encontrada: $($list.Title)" -f DarkGray
        }
    }
    Write-Log -Mensaje "Total de listas de OBIRAS encontradas: $($obirasCounter)"
    Write-Host "`nTotal de listas de OBIRAS encontradas: $($obirasCounter)`n" -f Green

    Write-Host ("─" * 50 + "`n") -f DarkGray
    foreach ($ListName in $obirasListNames) {
        Write-Log -Mensaje "Para la lista: '$ListName'"
        Write-Host "Para la lista: '$ListName'`n" -f DarkBlue
        
        Write-Log -Mensaje "Verificando la existencia de la lista '$ListName'..."
        Write-Host "Verificando la existencia de la lista '$ListName'..." -f Cyan

        $List = Get-PnPList -Identity $ListName -ErrorAction Stop
        Write-Log -Mensaje "La lista '$ListName' existe."
        Write-Host "La lista '$ListName' existe." -f Green
    
        Write-Log -Mensaje "Verificando la existencia de la columna '$DisplayName'..."
        Write-Host "`nVerificando la existencia de la columna '$DisplayName'..." -f Cyan
    
        $field = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
    
        if ($null -eq $field) {
            Write-Log -Mensaje "La columna '$DisplayName' no existe, creando..."
            Write-Host "La columna '$DisplayName' no existe, creando..." -f Green
            
            Add-PnPField -List $List `
                -DisplayName $DisplayName `
                -InternalName $InternalName `
                -Type Boolean
    
            Set-PnPField -List $List -Identity $InternalName -Values @{ DefaultFormula = "TRUE" }  | Out-Null
            
            Write-Log -Mensaje "La columna '$DisplayName' se agregó a la lista '$ListName' exitosamente."
            Write-Host "La columna '$DisplayName' se agregó a la lista '$ListName' exitosamente.`n" -f Green
        }
        else {
            Write-Log -Mensaje "La columna '$DisplayName' ya existe."
            Write-Host "La columna '$DisplayName' ya existe.`n" -f Green
        }
    
        Write-Log -Mensaje "Actualizando todos los ítems de la lista para establecer '$DisplayName' en 'Sí'..."
        Write-Host "Actualizando todos los ítems de la lista para establecer '$DisplayName' en 'Sí'...`n" -f Cyan
    
        $items = Get-PnPListItem -List $ListName -PageSize 1000
    
        foreach ($item in $items) {
            Set-PnPListItem -List $ListName -Identity $item.Id -Values @{ $InternalName = $true } | Out-Null
            Write-Log -Mensaje "Ítem ID $($item.Id) actualizado."
            Write-Host "Ítem ID $($item.Id) actualizado." -f Green
        }
    
        Write-Log -Mensaje "Todos los ítems de la lista '$ListName' fueron actualizados con '$DisplayName' en 'Sí'."
        Write-Host "`nTodos los ítems de la lista '$ListName' fueron actualizados con '$DisplayName' en 'Sí'.`n" -f Green
        Write-Host ("─" * 50 + "`n") -f DarkGray
    }

    Write-Log -Mensaje "Script ejecutado satisfactoriamente."
    Write-Host "Script ejecutado satisfactoriamente." -f Green

    $fin = Get-Date
    $duracion = $fin - $inicio
    $duracionFormateada = "{0:00}:{1:00}:{2:00}.{3:000}" -f $duracion.Hours, $duracion.Minutes, $duracion.Seconds, $duracion.Milliseconds

    Write-Log -Mensaje "Tiempo total de ejecución del script [hh:mm:ss]: $duracionFormateada"
    Write-Host "Tiempo total de ejecución del script [hh:mm:ss]: $duracionFormateada" -f Green
}
Catch {
    Write-Log -Nivel "ERROR" -Mensaje "Hubo un error: $($_.Exception.Message)"
    Write-Host -f Red "Hubo un error: " $_.Exception.Message
}
Finally {
    Disconnect-PnPOnline
    Write-Log -Mensaje "Conexión cerrada."
    Write-Host "Conexión cerrada." -f Magenta

    if ($global:stream) {
        $global:stream.Close()
    }
}
