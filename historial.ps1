$siteURL = "https://circo.sharepoint.com/sites/SHP-YPF-DEV/"

# ==========================================
# CONFIGURACION DE LOG
# ==========================================
$RutaLog = "C:/logs/"
$NombreBase = "FechaCierre"
$FechaHoraActual = Get-Date -Format "dd-MM-yyyy_HH-mm"
$NombreLog = "${NombreBase}_${FechaHoraActual}hs.log"
$RutaCompleta = Join-Path $RutaLog $NombreLog

# Formato solo para mostrar en consola/log
$formatoFechaLog = "dd/MM/yyyy HH:mm:ss"

function Write-Log {
    param (
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Nivel = "INFO",
        [Parameter(Mandatory)][string]$Mensaje
    )
    $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss.fff"
    $linea = "[$timestamp] [$Nivel] $Mensaje"
    try { $global:stream.WriteLine($linea) }
    catch { Write-Host "Error al escribir log: $_" -f Red }
}

# ----------------------------------------------------------
# Devuelve la version donde ocurrio la TRANSICION mas reciente
# al valor objetivo. Solo cuenta como transicion si en esa
# version el campo aparece en FieldValues y su valor cambia
# respecto del valor efectivo anterior.
# ----------------------------------------------------------
function Get-TransicionMasReciente {
    param(
        [Parameter(Mandatory)] $versions,
        [Parameter(Mandatory)] [string] $campo,
        [Parameter(Mandatory)] [string] $valorObjetivo
    )
    if (-not $versions -or $versions.Count -eq 0) { 
        return $null 
    }

    # Línea de tiempo de más vieja -> más nueva
    $asc = $versions | Sort-Object -Property Created

    $valorEfectivoActual = $null
    $ultimaTransicion = $null

    foreach ($v in $asc) {
        $tieneCampo = $false
        $valVersion = $null

        if ($v.FieldValues) {
            if ($v.FieldValues.ContainsKey($campo)) {
                $tieneCampo = $true
                $valVersion = $v.FieldValues[$campo]
            }
        }

        # Si el campo NO aparece en esta versión, se hereda el valor efectivo
        if (-not $tieneCampo) {
            $valVersion = $valorEfectivoActual
        }

        # Solo considero transición si el campo aparece y el valor CAMBIA
        if ($tieneCampo -and ($valVersion -ne $valorEfectivoActual)) {
            if ($valVersion -eq $valorObjetivo) {
                # registro esta transición; al final me quedo con la más reciente
                $ultimaTransicion = $v
            }
            $valorEfectivoActual = $valVersion
        }
        else {
            # inicialización: si nunca tuve valor y ahora conozco uno (heredado)
            if ($null -eq $valorEfectivoActual -and $null -ne $valVersion) {
                $valorEfectivoActual = $valVersion
            }
        }
    }

    return $ultimaTransicion
}

Try {
    $inicio = Get-Date

    if (-not (Test-Path -Path $RutaLog)) {
        New-Item -Path $RutaLog -ItemType Directory -Force | Out-Null
    }
    $global:stream = [System.IO.StreamWriter]::new($RutaCompleta, $true, [System.Text.Encoding]::UTF8)

    Write-Log -Mensaje "Inicio del script"
    Write-Host "Conectando a SharePoint Online..." -f Magenta
    Write-Log -Mensaje "Conectando a SharePoint Online..."
    $WarningPreference = "SilentlyContinue"
    Connect-PnPOnline -Url $SiteURL -UseWebLogin

    Write-Log -Mensaje "Obteniendo listas OBIRAS..."
    Write-Host "Obteniendo listas OBIRAS..." -f Cyan
    $lists = Get-PnPList | Where-Object { $_.Title -like "OBIRAS*" }
    $lists | ForEach-Object { 
        Write-Host " - '$($_.Title)'" -f DarkGray
        Write-Log -Mensaje " - '$($_.Title)'" 
    }

    $campoObjetivoEstadoGeneral = "EstadoGeneral"
    $valorObjetivoEstadoGeneral = "5. Cerrado"
    $campoFechaCierre = "FechaCierre"

    Write-Host "`n──────────────────────────────────────────────`n" -f DarkGray
    foreach ($list in $lists) {
        $tituloLista = $list.Title
        Write-Host "Procesando lista: '$tituloLista'" -f Cyan
        Write-Log -Mensaje "Procesando lista: '$tituloLista'"

        $items = Get-PnPListItem -List $tituloLista -PageSize 100
        Write-Host "Total de ítems: $($items.Count)" -f Gray
        Write-Log -Mensaje "Total de ítems: $($items.Count)"

        $encontrados = 0
        $actualizados = 0
        
        foreach ($item in $items) {
            if (-not $item[$campoFechaCierre]) {
                $versions = Get-PnPProperty -ClientObject $item -Property Versions
                $vTrans = Get-TransicionMasReciente -versions $versions -campo $campoObjetivoEstadoGeneral -valorObjetivo $valorObjetivoEstadoGeneral

                if (-not $vTrans) {
                    Write-Host "Ítem ID $($item.Id) no tiene '$campoObjetivoEstadoGeneral' = '$valorObjetivoEstadoGeneral', se omite." -f DarkGray
                    Write-Log -Mensaje "Ítem ID $($item.Id) no tiene '$campoObjetivoEstadoGeneral' = '$valorObjetivoEstadoGeneral', se omite."
                    continue
                }

                $encontrados++
                $fechaCierreLocal = $vTrans.Created.ToLocalTime()
                $fechaLog = $fechaCierreLocal.ToString($formatoFechaLog)

                Write-Host "=====================================" -f Yellow
                Write-Host "Ítem ID: $($item.Id)" -f White
                Write-Log -Mensaje "Ítem ID: $($item.Id)"
                Write-Host "'$campoObjetivoEstadoGeneral' = '$valorObjetivoEstadoGeneral' detectado en versión $($vTrans.VersionLabel)" -f Green
                Write-Log -Mensaje "'$campoObjetivoEstadoGeneral' = '$valorObjetivoEstadoGeneral' detectado en versión $($vTrans.VersionLabel)"
                Write-Host "Fecha detectada: $fechaLog" -f Gray
                Write-Log -Mensaje "Fecha detectada: $fechaLog"

                # Guardar DateTime (no string)
                Set-PnPListItem -List $tituloLista -Identity $item.Id -Values @{ $campoFechaCierre = $fechaCierreLocal } | Out-Null

                Write-Host "'$campoFechaCierre' seteada con éxito." -f Green
                Write-Log -Mensaje "Ítem $($item.Id): '$campoFechaCierre' seteada con éxito."
                Write-Host "=====================================" -f Yellow
                $actualizados++
            }
            else {
                Write-Host "Ítem ID $($item.Id) ya tiene '$campoFechaCierre', se omite." -f DarkGray
                Write-Log -Mensaje "Ítem $($item.Id): ya tiene '$campoFechaCierre', omitido."
            }
        }

        Write-Host ""
        Write-Host "────────── RESUMEN DE PROCESAMIENTO ──────────" -f DarkGray
        Write-Log -Mensaje "RESUMEN DE PROCESAMIENTO"
        Write-Host ("Lista: " + $tituloLista) -f Cyan
        Write-Log -Mensaje ("Lista: " + $tituloLista)
        Write-Host ("  - Ítems analizados: " + $items.Count)
        Write-Log -Mensaje ("  - Ítems analizados: " + $items.Count)
        Write-Host ("  - Ítems con '$campoObjetivoEstadoGeneral' = '$valorObjetivoEstadoGeneral' sin valor en '$campoFechaCierre': " + $encontrados)
        Write-Log -Mensaje ("  - Ítems con '$campoObjetivoEstadoGeneral' = '$valorObjetivoEstadoGeneral' sin valor en '$campoFechaCierre': " + $encontrados)
        Write-Host ("  - Ítems actualizados ('$campoFechaCierre' seteada): " + $actualizados)
        Write-Log -Mensaje ("  - Ítems actualizados ('$campoFechaCierre' seteada): " + $actualizados)
        $omitidos = $items.Count - $actualizados
        Write-Host ("  - Ítems omitidos (sin actualizar): " + $omitidos)
        Write-Log -Mensaje ("  - Ítems omitidos (sin actualizar): " + $omitidos)
        Write-Host "──────────────────────────────────────────────" -f DarkGray
        Write-Host ""
    }

    Write-Host "Script ejecutado satisfactoriamente." -f Green
    Write-Log -Mensaje "Script ejecutado satisfactoriamente."

    $fin = Get-Date
    $duracion = $fin - $inicio
    $duracionFormateada = "{0:00}:{1:00}:{2:00}.{3:000}" -f $duracion.Hours, $duracion.Minutes, $duracion.Seconds, $duracion.Milliseconds
    Write-Host "Tiempo total [hh:mm:ss]: $duracionFormateada" -f Green
    Write-Log -Mensaje "Tiempo total [hh:mm:ss]: $duracionFormateada"
}
Catch {
    Write-Log -Nivel "ERROR" -Mensaje "Hubo un error: $($_.Exception.Message)"
    Write-Host -f Red "Hubo un error: " $_.Exception.Message
}
Finally {
    Disconnect-PnPOnline
    Write-Host "Conexión cerrada." -f Magenta
    Write-Log -Mensaje "Conexión cerrada."
    Write-Log -Mensaje "Fin del script."
    if ($global:stream) { 
        $global:stream.Close() 
    }
}
