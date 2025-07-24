# ==========================================
# CONFIGURACION DE LOG
# ==========================================
$RutaLog = "C:/Users/PC/Desktop/logs"
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
# FUNCION PARA VALIDAR EXISTENCIA DE LISTA Y CAMPOS
# ==========================================
function ValidarListaYCampos {
    param (
        [string]$NombreDeLista,
        [string[]]$NombresInternosDeCampos
    )

    try {
        Get-PnPList -Identity $NombreDeLista -ThrowExceptionIfListNotFound | Out-Null
    } catch {
        throw "La lista '$NombreDeLista' no existe en el sitio."
    }

    $camposExistentes = Get-PnPField -List $NombreDeLista | Select-Object -ExpandProperty InternalName

    foreach ($campo in $NombresInternosDeCampos) {
        if ($campo -notin $camposExistentes) {
            throw "El campo '$campo' no existe en la lista '$NombreDeLista'."
        }
    }

    Write-Host "Validación exitosa para la lista '$NombreDeLista'" -ForegroundColor Green
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

    # VARIABLES GEERALES
    $tamPagina = 500

    # VARIABLES: Secciones
    $SeccionesNombreDeLista = "Secciones"
    $SeccionesInternalNameID = "ID"
    $SeccionesInternalNameSeccion = "Title"
    $SeccionesInternalNameArea = "FdsArea"

    # VARIABLES: Productos
    $ProductosNombreDeLista = "Productos"
    $ProductosInternalNameID = "ID"
    $ProductosInternalNameProducto = "Title"
    $ProductosInternalNameSeccion = "FdsSeccion"
    $ProductosInternalNameArea = "FdsArea"

    # Validación de existencia: lista y campos Secciones
    ValidarListaYCampos -NombreDeLista $SeccionesNombreDeLista -NombresInternosDeCampos @(
        $SeccionesInternalNameID,
        $SeccionesInternalNameSeccion,
        $SeccionesInternalNameArea
    )

    # Validación de existencia: lista y campos Productos
    ValidarListaYCampos -NombreDeLista $ProductosNombreDeLista -NombresInternosDeCampos @(
        $ProductosInternalNameID,
        $ProductosInternalNameProducto,
        $ProductosInternalNameSeccion,
        $ProductosInternalNameArea
    )


    # Ejecutar query: llamada bloqueante al sitio de SharePoint, hasta que no traiga TODOS los registros, trabajando con paginado de $tamPagina, NO CONTINUA EJECUCIÓN
    $itemsDeSecciones = Get-PnPListItem -List $SeccionesNombreDeLista `
        -Fields $SeccionesInternalNameID, $SeccionesInternalNameSeccion, $SeccionesInternalNameArea `
        -PageSize $tamPagina

    # Mostrar resultados
    foreach ($seccion in $itemsDeSecciones) {
        $id = $seccion[$SeccionesInternalNameID]
        $titulo = $seccion[$SeccionesInternalNameSeccion]

        # Acceder al campo lookup como FieldLookupValue
        $area = $seccion.FieldValues[$SeccionesInternalNameArea]

        # Puede que venga null o como string, verificar tipo
        if ($area -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
            $areaId = $area.LookupId
            $areaNombre = $area.LookupValue
        } else {
            $areaId = ""
            $areaNombre = $area
        }

        Write-Host "ID: $id - Titulo: $titulo - Area: $areaNombre (ID: $areaId)" -ForegroundColor DarkGray
    }

$camlQuerySeccion = @"
<View>
  <Query>
    <OrderBy>
      <FieldRef Name="$SeccionesInternalNameSeccion" Ascending='FALSE' />
    </OrderBy>
  </Query>
</View>
"@

    # Obtener productos
    # $itemsDeProductos = Get-PnPListItem -List $ProductosNombreDeLista `
    #     -Fields $ProductosInternalNameID, $ProductosInternalNameProducto, $ProductosInternalNameSeccion, $ProductosInternalNameArea `
    #     -PageSize $tamPagina
    $itemsDeProductos = Get-PnPListItem -List $ProductosNombreDeLista `
    -Query $camlQuerySeccion -PageSize $tamPagina

    # Mostrar resultados
    foreach ($producto in $itemsDeProductos) {
        $id = $producto[$ProductosInternalNameID]
        $titulo = $producto[$ProductosInternalNameProducto]

        # Campo Seccion (lookup)
        $seccion = $producto.FieldValues[$ProductosInternalNameSeccion]
        if ($seccion -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
            $seccionId = $seccion.LookupId
            $seccionNombre = $seccion.LookupValue
        } else {
            $seccionId = ""
            $seccionNombre = $seccion
        }

        # Campo Area (lookup)
        $area = $producto.FieldValues[$ProductosInternalNameArea]
        if ($area -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
            $areaId = $area.LookupId
            $areaNombre = $area.LookupValue
        } else {
            $areaId = ""
            $areaNombre = $area
        }

        Write-Host "ID: $id - Titulo: $titulo - Seccion: $seccionNombre (ID: $seccionId) - Area: $areaNombre (ID: $areaId)" -ForegroundColor DarkGray
    }





    # LÓGICA DE LA ACCIÓN PRINCIPAL
    # for ($i = 1; $i -le 5; $i++) {
    #     $mensaje = "Registro de prueba $i"
    #     Write-Host "$mensaje" -ForegroundColor DarkGray
    #     Write-Log -Nivel "INFO" -Mensaje $mensaje
    # }

    # ==========================================
    # FIN DE ACCIÓN PRINCIPAL
    # ==========================================
    Write-Host "Script ejecutado satisfactoriamente." -ForegroundColor Green
    Write-Log -Nivel "INFO" -Mensaje "Script ejecutado satisfactoriamente"

    $fin = Get-Date
    $duracion = $fin - $inicio
    $duracionFormateada = "{0:00}:{1:00}:{2:00}.{3:000}" -f $duracion.Hours, $duracion.Minutes, $duracion.Seconds, $duracion.Milliseconds

    Write-Log -Nivel "INFO" -Mensaje "Tiempo total de ejecución del script [hh:mm:ss]: $duracionFormateada"
    Write-Host "Tiempo total de ejecución del script [hh:mm:ss]: $duracionFormateada" -ForegroundColor Green
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
