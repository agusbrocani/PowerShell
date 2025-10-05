# param (
#     [Parameter(Mandatory)][string]$SiteURL
# )

$SiteURL = "https://circo.sharepoint.com/sites/SHP-YPF-DEV/"

# ==========================================
# CONFIGURACION DE LOG
# ==========================================
$RutaLog = "C:/logs/"
$NombreBase = "CrearLookupApuntandoRegistros"
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

$oldInternalName = "Area"
$InternalName = "Bloque"
$displayName = "Bloque"
$PADFieldInternalName = "PADLocacion"
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
        
        Write-Log -Mensaje "Verificando la existencia de la lista..."
        Write-Host "Verificando la existencia de la lista..." -f Cyan

        $List = Get-PnPList -Identity $ListName -ErrorAction Stop
        Write-Host "La lista existe." -f Green
        Write-Log -Mensaje "La lista existe."
    
        $fieldArea = Get-PnPField -List $list -Identity $oldInternalName -ErrorAction SilentlyContinue

        Write-Log -Mensaje "Eliminando el campo '$oldInternalName' de la lista..."
        Write-Host "`nEliminando el campo '$oldInternalName' de la lista..." -f Cyan
        if ($fieldArea) {
            Remove-PnPField -List $list -Identity $oldInternalName -Force
            Write-Log -Mensaje "Campo '$oldInternalName' eliminado correctamente."
            Write-Host "Campo '$oldInternalName' eliminado correctamente." -f Green
        }
        else {
            Write-Log -Nivel "WARNING" -Mensaje "Campo '$oldInternalName' no existe en la lista."
            Write-Host "Campo '$oldInternalName' no existe en la lista." -f Yellow
        }

        $locacionesList = Get-PnPList -Identity "Locaciones"
        $bloqueField = Get-PnPField -List $list -Identity $internalName -ErrorAction SilentlyContinue

        Write-Host "`nCreando campo lookup '$displayName' en la lista..." -f Cyan
        Write-Log -Mensaje "Creando campo lookup '$displayName' en la lista..."
        if (-not $bloqueField) {
            $areaInternalName = "AREA"
            # Definir el schema del campo lookup
            $fieldSchema = @"
<Field 
    Type='Lookup'
    DisplayName='$displayName'
    StaticName='$internalName'
    Name='$internalName'
    List='{$($locacionesList.Id)}'
    ShowField='$areaInternalName'
    />
"@

            Add-PnPFieldFromXml -List $list -FieldXml $fieldSchema | Out-Null
            Write-Log -Mensaje "Campo lookup '$displayName' creado correctamente apuntando al campo '$areaInternalName' de la lista '$($locacionesList.Title)'."
            Write-Host "Campo lookup '$displayName' creado correctamente apuntando al campo '$areaInternalName' de la lista '$($locacionesList.Title)'." -f Green
        }
        else {
            Write-Log -Nivel "WARNING" -Mensaje "El campo '$displayName' ya existe en la lista. No se creó de nuevo."
            Write-Host "El campo '$displayName' ya existe en la lista. No se creó de nuevo." -f Yellow
        }

        $items = Get-PnPListItem -List $list -Fields "ID", $PADFieldInternalName

        Write-Host "`nActualizando valor del campo lookup '$displayName' con LookupId del campo '$PADFieldInternalName' para todo ítem de la lista..." -f Cyan
        Write-Log -Mensaje "Actualizando valor del campo lookup '$displayName' con LookupId del campo '$PADFieldInternalName' para todo ítem de la lista..."
        foreach ($item in $items) {
            $idItem = $item["ID"]
            $padId = $item[$PADFieldInternalName]?.LookupId

            if ($padId) {
                Set-PnPListItem -List $list -Identity $idItem -Values @{
                    $InternalName = $padId
                } | Out-Null

                Write-Log -Mensaje "Campo '$displayName' del ítem '$idItem' actualizado con LookupId '$($padId)'."
                Write-Host "Campo '$displayName' del ítem '$idItem' actualizado con LookupId '$($padId)'." -f Green
            }
            else {
                Write-Log -Mensaje "Ítem '$idItem' no tiene valor en el campo '$PADFieldInternalName' de la lista. No se actualizó."
                Write-Host "Ítem '$idItem' no tiene valor en el campo '$PADFieldInternalName' de la lista. No se actualizó." -f Yellow
            }
        }

        Write-Log -Mensaje "Fin del proceso para la lista: '$ListName'"
        Write-Host "`nFin del proceso para la lista: '$ListName'`n" -f DarkBlue
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
