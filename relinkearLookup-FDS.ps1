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

$camlQuerySecciones = @"
<View>
  <Query>
    <OrderBy>
      <FieldRef Name="$SeccionesInternalNameSeccion" Ascending='TRUE' />
    </OrderBy>
  </Query>
</View>
"@

    # Ejecutar query: llamada bloqueante al sitio de SharePoint, hasta que no traiga TODOS los registros, trabajando con paginado de $tamPagina, NO CONTINUA EJECUCIÓN
        $itemsDeSecciones = Get-PnPListItem -List $SeccionesNombreDeLista `
        -Query $camlQuerySecciones -PageSize $tamPagina

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

Write-Host "ESTA ES LA LOGICA PARA MANTENER REGISTROS DUPLICADOS" -ForegroundColor Red

# Case sensitive (usa StringComparer.Ordinal)
$contadorPorTitulo = New-Object "System.Collections.Generic.Dictionary[String, Int32]" ([StringComparer]::Ordinal)
$seccionesDuplicadas = New-Object "System.Collections.Generic.Dictionary[Int32, Object]"

# 1. Contar títulos y guardar duplicados (segunda aparición en adelante)
foreach ($seccion in $itemsDeSecciones) {
    $titulo = $seccion[$SeccionesInternalNameSeccion]

    if ($contadorPorTitulo.ContainsKey($titulo)) {
        $contadorPorTitulo[$titulo] += 1
        $seccionesDuplicadas[$seccion[$SeccionesInternalNameID]] = $seccion
    } else {
        $contadorPorTitulo[$titulo] = 1
    }
}

# 2. Agregar la primera aparición de cada título duplicado
foreach ($titulo in $contadorPorTitulo.Keys) {
    if ($contadorPorTitulo[$titulo] -gt 1) {
        $primeraCoincidencia = $itemsDeSecciones | Where-Object { $_[$SeccionesInternalNameSeccion] -eq $titulo } | Select-Object -First 1
        $seccionesDuplicadas[$primeraCoincidencia[$SeccionesInternalNameID]] = $primeraCoincidencia
    }
}

# 3. Mostrar duplicados ordenados por Título
$seccionesDuplicadas.Values | Sort-Object { $_[$SeccionesInternalNameSeccion] } | ForEach-Object {
    $id = $_[$SeccionesInternalNameID]
    $titulo = $_[$SeccionesInternalNameSeccion]

    $area = $_[$SeccionesInternalNameArea]
    if ($area -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
        $areaId = $area.LookupId
        $areaNombre = $area.LookupValue
    } else {
        $areaId = ""
        $areaNombre = $area
    }

    Write-Host "ID: $id - Titulo: $titulo - Area: $areaNombre (ID: $areaId)" -ForegroundColor Yellow
}


    

$camlQueryProductos = @"
<View>
  <Query>
    <OrderBy>
      <FieldRef Name="$SeccionesInternalNameSeccion" Ascending='FALSE' />
    </OrderBy>
  </Query>
</View>
"@

    $itemsDeProductos = Get-PnPListItem -List $ProductosNombreDeLista `
    -Query $camlQueryProductos -PageSize $tamPagina

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

Write-Host "ESTA ES LA LOGICA PARA REAPUNTAR IDs LOOKUP FdsSeccion en Productos" -ForegroundColor Red
Write-Log -Nivel "INFO" -Mensaje "Iniciando lógica para reapuntar IDs de FdsSeccion en Productos"

foreach ($producto in $itemsDeProductos) {
    $productoId = $producto[$ProductosInternalNameID]
    $productoTitulo = $producto[$ProductosInternalNameProducto]

    $productoSeccion = $producto.FieldValues[$ProductosInternalNameSeccion]
    $productoSeccionId = if ($productoSeccion -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $productoSeccion.LookupId } else { "" }
    $productoSeccionNombre = if ($productoSeccion -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $productoSeccion.LookupValue } else { $productoSeccion }

    $productoArea = $producto.FieldValues[$ProductosInternalNameArea]
    $productoAreaNombre = if ($productoArea -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $productoArea.LookupValue } else { $productoArea }

    # Buscar en seccionesDuplicadas una coincidencia por Título Y Área (case sensitive)
    $seccionCorrecta = $seccionesDuplicadas.Values | Where-Object {
        $_[$SeccionesInternalNameSeccion] -eq $productoSeccionNombre -and (
            ($_.FieldValues[$SeccionesInternalNameArea]?.LookupValue) -eq $productoAreaNombre
        )
    } | Select-Object -First 1

    if ($seccionCorrecta) {
        $seccionCorrectaId = $seccionCorrecta[$SeccionesInternalNameID]

        if ($productoSeccionId -ne $seccionCorrectaId) {
            Write-Host "Corrigiendo producto ID: $productoId - '$productoTitulo'" -ForegroundColor Cyan
            Write-Host " → De Sección ID: $productoSeccionId a ID: $seccionCorrectaId" -ForegroundColor Yellow

            Write-Log -Nivel "INFO" -Mensaje "Corrigiendo Producto ID: $productoId - '$productoTitulo' | Sección: '$productoSeccionNombre' | Área: '$productoAreaNombre' | De ID: $productoSeccionId → A ID: $seccionCorrectaId"

            # Actualizar el lookup con el nuevo ID
            try {
                Set-PnPListItem -List $ProductosNombreDeLista -Identity $productoId -Values @{
                    $ProductosInternalNameSeccion = $seccionCorrectaId
                }
            } catch {
                Write-Host "Error al actualizar producto ID: $productoId" -ForegroundColor Red
                Write-Log -Nivel "ERROR" -Mensaje "Error al actualizar Producto ID: $productoId - '$productoTitulo' → $_"
            }
        }
    }
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