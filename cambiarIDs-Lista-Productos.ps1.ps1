# ==========================================
# CONFIGURACION DE LOG
# ==========================================
$RutaLog = "C:/Users/PC/Desktop/logs"
$NombreBase = "LISTA_PRODUCTOS"
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
        throw "La lista '$NombreDeLista' no existe en el sitio"
    }

    $camposExistentes = Get-PnPField -List $NombreDeLista | Select-Object -ExpandProperty InternalName

    foreach ($campo in $NombresInternosDeCampos) {
        if ($campo -notin $camposExistentes) {
            throw "El campo '$campo' no existe en la lista '$NombreDeLista'"
        }
    }

    Write-Host "Validacion exitosa para la lista '$NombreDeLista'" -ForegroundColor Green
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
    # INICIO DE VALIDAR EXISTENCIA DE ESTRUCTURAS Y CAMPOS
    # ==========================================
    Write-Host "`nInicio de validaciones de existencia de estructuras y campos..." -ForegroundColor Cyan
    Write-Log -Nivel "INFO" -Mensaje "Inicio de validaciones de existencia de estructuras y campos..."

    # VARIABLES GENERALES
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

    # Validacion de existencia: lista y campos Secciones
    ValidarListaYCampos -NombreDeLista $SeccionesNombreDeLista -NombresInternosDeCampos @(
        $SeccionesInternalNameID,
        $SeccionesInternalNameSeccion,
        $SeccionesInternalNameArea
    )
    Write-Log -Mensaje "Validaciones exitosas para la lista '$SeccionesNombreDeLista'"

    # Validacion de existencia: lista y campos Productos
    ValidarListaYCampos -NombreDeLista $ProductosNombreDeLista -NombresInternosDeCampos @(
        $ProductosInternalNameID,
        $ProductosInternalNameProducto,
        $ProductosInternalNameSeccion,
        $ProductosInternalNameArea
    )
    Write-Log -Mensaje "Validaciones exitosas para la lista '$ProductosNombreDeLista'"
    # ==========================================
    # FIN DE VALIDAR EXISTENCIA DE ESTRUCTURA Y CAMPOS
    # ==========================================

    # ==========================================
    # INICIO DE OBTENCION DE REGISTROS EN LISTA SECCION
    # ==========================================
    Write-Host "`nObteniendo registros de la lista '$SeccionesNombreDeLista'..." -ForegroundColor Cyan
    Write-Log -Nivel "INFO" -Mensaje "Obteniendo registros de la lista '$SeccionesNombreDeLista'..."

$camlQuerySecciones = @"
<View>
  <Query>
    <OrderBy>
      <FieldRef Name="$SeccionesInternalNameSeccion" Ascending='TRUE' />
    </OrderBy>
  </Query>
</View>
"@

    # Ejecutar query: llamada bloqueante al sitio de SharePoint, hasta que no traiga TODOS los registros, trabajando con paginado de $tamPagina, NO CONTINUA EJECUCION
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

    Write-Host "Registros de la lista '$SeccionesNombreDeLista' obtenidos con exito" -ForegroundColor Green
    Write-Log -Mensaje "Registros de la lista '$SeccionesNombreDeLista' obtenidos con exito"
    # ==========================================
    # FIN DE OBTENCION DE REGISTROS EN LISTA SECCION
    # ==========================================

    # ==========================================
    # INICIO DE OBTENER REGISTROS CON Seccion DUPLICADA EN LA LISTA DE SECCIONES
    # ==========================================
    Write-Host "`nObteniendo registros con el valor del campo 'Seccion' duplicado en la lista '$SeccionesNombreDeLista'..." -ForegroundColor Cyan
    Write-Log -Mensaje "Obteniendo registros con el valor del campo 'Seccion' duplicado en la lista '$SeccionesNombreDeLista'..."

    # Case sensitive (usa StringComparer.Ordinal)
    $contadorPorTitulo = New-Object "System.Collections.Generic.Dictionary[String, Int32]" ([StringComparer]::Ordinal)
    $seccionesDuplicadas = New-Object "System.Collections.Generic.Dictionary[Int32, Object]"

    # 1. Contar titulos y guardar duplicados (segunda aparicion en adelante)
    foreach ($seccion in $itemsDeSecciones) {
        $titulo = $seccion[$SeccionesInternalNameSeccion]

        if ($contadorPorTitulo.ContainsKey($titulo)) {
            $contadorPorTitulo[$titulo] += 1
            $seccionesDuplicadas[$seccion[$SeccionesInternalNameID]] = $seccion
        } else {
            $contadorPorTitulo[$titulo] = 1
        }
    }

    # 2. Agregar la primera aparicion de cada titulo duplicado
    foreach ($titulo in $contadorPorTitulo.Keys) {
        if ($contadorPorTitulo[$titulo] -gt 1) {
            $primeraCoincidencia = $itemsDeSecciones | Where-Object { $_[$SeccionesInternalNameSeccion] -eq $titulo } | Select-Object -First 1
            $seccionesDuplicadas[$primeraCoincidencia[$SeccionesInternalNameID]] = $primeraCoincidencia
        }
    }

    # 3. Mostrar duplicados ordenados por Titulo
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

        Write-Host "ID: $id - Titulo: $titulo - Area: $areaNombre (ID: $areaId)" -ForegroundColor DarkGray
    }

    Write-Host "Registros con el valor del campo 'Seccion' duplicado en la lista '$SeccionesNombreDeLista' obtenidos con exito" -ForegroundColor Green
    Write-Log -Mensaje "Registros con el valor del campo 'Seccion' duplicado en la lista '$SeccionesNombreDeLista' obtenidos con exito"
    # ==========================================
    # FIN DE OBTENER REGISTROS CON Seccion DUPLICADA EN LA LISTA DE SECCIONES
    # ==========================================

    # ==========================================
    # INICIO DE OBTENCION DE REGISTROS EN LISTA PRODUCTOS
    # ==========================================
    Write-Host "`nObteniendo registros de la lista '$ProductosNombreDeLista'..." -ForegroundColor Cyan
    Write-Log -Nivel "INFO" -Mensaje "Obteniendo registros de la lista '$ProductosNombreDeLista'..."
$camlQueryProductos = @"
<View>
  <Query>
    <OrderBy>
      <FieldRef Name="$ProductosInternalNameProducto" Ascending='FALSE' />
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

    Write-Host "Registros de la lista '$ProductosNombreDeLista' obtenidos con exito" -ForegroundColor Green
    Write-Log -Mensaje "Registros de la lista '$ProductosNombreDeLista' obtenidos con exito"
    # ==========================================
    # FIN DE OBTENCION DE REGISTROS EN LISTA PRODUCTOS
    # ==========================================

    # ==========================================
    # INICIO DE LOGICA PARA REAPUNTAR IDs LOOKUP PARA EL CAMPO FdsSeccion EN LA LISTA PRODUCTOS
    # ==========================================
    Write-Host "`nIniciando logica para reapuntar IDs Lookup erroneos del campo 'Seccion' en la lista '$ProductosNombreDeLista'..." -ForegroundColor Cyan
    Write-Log -Nivel "INFO" -Mensaje "Iniciando logica para reapuntar IDs Lookup erroneos del campo 'Seccion' en la lista '$ProductosNombreDeLista'..."

    foreach ($producto in $itemsDeProductos) {
        $productoId = $producto[$ProductosInternalNameID]
        $productoTitulo = $producto[$ProductosInternalNameProducto]

        $productoSeccion = $producto.FieldValues[$ProductosInternalNameSeccion]
        $productoSeccionId = if ($productoSeccion -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $productoSeccion.LookupId } else { "" }
        $productoSeccionNombre = if ($productoSeccion -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $productoSeccion.LookupValue } else { $productoSeccion }

        $productoArea = $producto.FieldValues[$ProductosInternalNameArea]
        $productoAreaNombre = if ($productoArea -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $productoArea.LookupValue } else { $productoArea }

        # Buscar en seccionesDuplicadas una coincidencia por Titulo Y Area (case sensitive)
        $seccionCorrecta = $seccionesDuplicadas.Values | Where-Object {
            $_[$SeccionesInternalNameSeccion] -eq $productoSeccionNombre -and (
                ($_.FieldValues[$SeccionesInternalNameArea]?.LookupValue) -eq $productoAreaNombre
            )
        } | Select-Object -First 1  # Si llegara a haber mas de 1 match, solo tomo el primero

        if ($seccionCorrecta) {
            $seccionCorrectaId = $seccionCorrecta[$SeccionesInternalNameID]

            if ($productoSeccionId -ne $seccionCorrectaId) {
                Write-Host "Corrigiendo producto ID: $productoId - '$productoTitulo'" -ForegroundColor DarkBlue
                Write-Host " -> De Seccion ID: $productoSeccionId a ID: $seccionCorrectaId" -ForegroundColor DarkBlue

                Write-Log -Nivel "INFO" -Mensaje "Corrigiendo ID: $productoId - Producto: '$productoTitulo' | Seccion: '$productoSeccionNombre' | Area: '$productoAreaNombre' | De ID: $productoSeccionId -> A ID: $seccionCorrectaId"

                # Actualizar el lookup con el nuevo ID
                try {
                    Set-PnPListItem -List $ProductosNombreDeLista -Identity $productoId -Values @{
                        $ProductosInternalNameSeccion = $seccionCorrectaId
                    }
                } catch {
                    Write-Host "Error al actualizar producto ID: $productoId" -ForegroundColor Red
                    Write-Log -Nivel "ERROR" -Mensaje "Error al actualizar ID: $productoId - Producto: '$productoTitulo' -> $_"
                }
            }
        }
    }

    Write-Host "Logica para reapuntar IDs Lookup erroneos del campo 'Seccion' en la lista '$ProductosNombreDeLista' finalizada con exito" -ForegroundColor Green
    Write-Log -Nivel "INFO" -Mensaje "Logica para reapuntar IDs Lookup erroneos del campo 'Seccion' en la lista '$ProductosNombreDeLista' finalizada con exito"
    # ==========================================
    # FIN DE LOGICA PARA REAPUNTAR IDs LOOKUP PARA EL CAMPO FdsSeccion EN LA LISTA PRODUCTOS
    # ==========================================

    Write-Host "`nEstado de ejecucion:" -ForegroundColor Cyan
    Write-Host "Script ejecutado satisfactoriamente." -ForegroundColor Green
    Write-Log -Nivel "INFO" -Mensaje "Script ejecutado satisfactoriamente"

    $fin = Get-Date
    $duracion = $fin - $inicio
    $duracionFormateada = "{0:00}:{1:00}:{2:00}.{3:000}" -f $duracion.Hours, $duracion.Minutes, $duracion.Seconds, $duracion.Milliseconds

    Write-Log -Nivel "INFO" -Mensaje "Tiempo total de ejecucion del script [hh:mm:ss]: $duracionFormateada"
    Write-Host "Tiempo total de ejecucion del script [hh:mm:ss]: $duracionFormateada" -ForegroundColor Green
}
catch {
    Write-Host "¡Error! $($_.Exception.Message)" -ForegroundColor Red
    Write-Log -Nivel "ERROR" -Mensaje $_.Exception.Message
}
finally {
    Disconnect-PnPOnline
    Write-Host "`nConexion cerrada." -ForegroundColor Magenta
    Write-Log -Nivel "INFO" -Mensaje "Conexion cerrada"

    if ($global:stream) {
        $global:stream.Close()
    }
}

# Uso apropiado de colores
    # Magenta   -> recursos
    # Cyan      -> acciones
    # Green     -> acciones completadas con exito
    # Red       -> errores
    # Yellow    -> warnings
    # DarkGrey  -> uso comun

# ESCRIBIR en archivo log:
    # Write-Log -Nivel "ERROR" -Mensaje $mensajeAEscribir
    # -Nivel    [OPCIONAL]. Valores posibles: "INFO", "WARNING", "ERROR". Por defecto es "INFO". 
    # -Mensaje  [OBLIGATORIO]
