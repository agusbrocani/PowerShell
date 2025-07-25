# ==========================================
# CONFIGURACION DE LOG
# ==========================================
$RutaLog = "C:/Users/PC/Desktop/logs"
$NombreBase = "BIBLIOTECA_FICHAS_DE_INTERVENCION_DE_EMERGENCIA"
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
# FUNCION PARA VALIDAR EXISTENCIA DE ESTRUCTURA Y CAMPOS
# ==========================================
function validarEstructuraYCampos {
    param (
        [string]$NombreDeEstructura,
        [string[]]$NombresInternosDeCampos
    )

    try {
        Get-PnPList -Identity $NombreDeEstructura -ThrowExceptionIfListNotFound | Out-Null
    } catch {
        throw "La estructura '$NombreDeEstructura' no existe en el sitio"
    }

    $camposExistentes = Get-PnPField -List $NombreDeEstructura | Select-Object -ExpandProperty InternalName

    foreach ($campo in $NombresInternosDeCampos) {
        if ($campo -notin $camposExistentes) {
            throw "El campo '$campo' no existe en la estructura '$NombreDeEstructura'"
        }
    }

    Write-Host "Validacion exitosa para la estructura '$NombreDeEstructura'" -ForegroundColor Green
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

    # VARIABLES: Fichas De Intervencion De Emergencia
    $NombreDeBiblioteca = "Fichas De Intervencion De Emergencia"
    $BibliotecaInternalNameID = "ID"
    $BibliotecaInternalNameNombre = "Title"
    $BibliotecaInternalNameSeccion = "FdsSeccion"
    $BibliotecaInternalNameArea = "FdsArea"

    # Validacion de existencia: lista y campos Secciones
    validarEstructuraYCampos -NombreDeEstructura $SeccionesNombreDeLista -NombresInternosDeCampos @(
        $SeccionesInternalNameID,
        $SeccionesInternalNameSeccion,
        $SeccionesInternalNameArea
    )
    Write-Log -Mensaje "Validaciones exitosas para la lista '$SeccionesNombreDeLista'"

    # Validacion de existencia: biblioteca y campos
    validarEstructuraYCampos -NombreDeEstructura $NombreDeBiblioteca -NombresInternosDeCampos @(
        $BibliotecaInternalNameID,
        $BibliotecaInternalNameNombre,
        $BibliotecaInternalNameSeccion,
        $BibliotecaInternalNameArea
    )
    Write-Log -Mensaje "Validaciones exitosas para la biblioteca '$NombreDeBiblioteca'"
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
    # INICIO DE OBTENCION DE REGISTROS EN BIBLIOTECA
    # ==========================================
    Write-Host "`nObteniendo registros de la biblioteca '$NombreDeBiblioteca'..." -ForegroundColor Cyan
    Write-Log -Nivel "INFO" -Mensaje "Obteniendo registros de la biblioteca '$NombreDeBiblioteca'..."
$camlQueryProductos = @"
<View>
  <Query>
    <OrderBy>
      <FieldRef Name="$BibliotecaInternalNameNombre" Ascending='FALSE' />
    </OrderBy>
  </Query>
</View>
"@

    $itemsDeBiblioteca = Get-PnPListItem -List $NombreDeBiblioteca `
    -Query $camlQueryProductos -PageSize $tamPagina

    # Mostrar resultados
    foreach ($biblioteca in $itemsDeBiblioteca) {
        $id = $biblioteca[$BibliotecaInternalNameID]
        $titulo = $biblioteca[$BibliotecaInternalNameNombre]

        # Campo Seccion (lookup)
        $seccion = $biblioteca.FieldValues[$BibliotecaInternalNameSeccion]
        if ($seccion -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
            $seccionId = $seccion.LookupId
            $seccionNombre = $seccion.LookupValue
        } else {
            $seccionId = ""
            $seccionNombre = $seccion
        }

        # Campo Area (lookup)
        $area = $biblioteca.FieldValues[$BibliotecaInternalNameArea]
        if ($area -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
            $areaId = $area.LookupId
            $areaNombre = $area.LookupValue
        } else {
            $areaId = ""
            $areaNombre = $area
        }

        Write-Host "ID: $id - Titulo: $titulo - Seccion: $seccionNombre (ID: $seccionId) - Area: $areaNombre (ID: $areaId)" -ForegroundColor DarkGray
    }

    Write-Host "Registros de la biblioteca '$NombreDeBiblioteca' obtenidos con exito" -ForegroundColor Green
    Write-Log -Mensaje "Registros de la biblioteca '$NombreDeBiblioteca' obtenidos con exito"
    # ==========================================
    # FIN DE OBTENCION DE REGISTROS EN BIBLIOTECA
    # ==========================================

    # ==========================================
    # INICIO DE LOGICA PARA REAPUNTAR IDs LOOKUP PARA EL CAMPO FdsSeccion EN LA BIBLIOTECA
    # ==========================================
    Write-Host "`nIniciando logica para reapuntar IDs Lookup erroneos del campo 'Seccion' en la biblioteca '$NombreDeBiblioteca'..." -ForegroundColor Cyan
    Write-Log -Nivel "INFO" -Mensaje "Iniciando logica para reapuntar IDs Lookup erroneos del campo 'Seccion' en la biblioteca '$NombreDeBiblioteca'..."

    foreach ($biblioteca in $itemsDeBiblioteca) {
        $bibliotecaId = $biblioteca[$BibliotecaInternalNameID]
        $bibliotecaTitulo = $biblioteca[$BibliotecaInternalNameNombre]

        $bibliotecaSeccion = $biblioteca.FieldValues[$BibliotecaInternalNameSeccion]
        $bibliotecaSeccionId = if ($bibliotecaSeccion -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $bibliotecaSeccion.LookupId } else { "" }
        $bibliotecaSeccionNombre = if ($bibliotecaSeccion -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $bibliotecaSeccion.LookupValue } else { $bibliotecaSeccion }

        $bibliotecaArea = $biblioteca.FieldValues[$BibliotecaInternalNameArea]
        $bibliotecaAreaNombre = if ($bibliotecaArea -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $bibliotecaArea.LookupValue } else { $bibliotecaArea }

        # Buscar en seccionesDuplicadas una coincidencia por Titulo Y Area (case sensitive)
        $seccionCorrecta = $seccionesDuplicadas.Values | Where-Object {
            $_[$SeccionesInternalNameSeccion] -eq $bibliotecaSeccionNombre -and (
                ($_.FieldValues[$SeccionesInternalNameArea]?.LookupValue) -eq $bibliotecaAreaNombre
            )
        } | Select-Object -First 1  # Si llegara a haber mas de 1 match, solo tomo el primero

        if ($seccionCorrecta) {
            $seccionCorrectaId = $seccionCorrecta[$SeccionesInternalNameID]

            if ($bibliotecaSeccionId -ne $seccionCorrectaId) {
                Write-Host "Corrigiendo biblioteca ID: $bibliotecaId - '$bibliotecaTitulo'" -ForegroundColor DarkBlue
                Write-Host " -> De Seccion ID: $bibliotecaSeccionId a ID: $seccionCorrectaId" -ForegroundColor DarkBlue

                Write-Log -Nivel "INFO" -Mensaje "Corrigiendo ID: $bibliotecaId - Producto: '$bibliotecaTitulo' | Seccion: '$bibliotecaSeccionNombre' | Area: '$bibliotecaAreaNombre' | De ID: $bibliotecaSeccionId -> A ID: $seccionCorrectaId"

                # Actualizar el lookup con el nuevo ID
                try {
                    Set-PnPListItem -List $NombreDeBiblioteca -Identity $bibliotecaId -Values @{
                        $BibliotecaInternalNameSeccion = $seccionCorrectaId
                    }
                } catch {
                    Write-Host "Error al actualizar biblioteca ID: $bibliotecaId" -ForegroundColor Red
                    Write-Log -Nivel "ERROR" -Mensaje "Error al actualizar ID: $bibliotecaId - Producto: '$bibliotecaTitulo' -> $_"
                }
            }
        }
    }

    Write-Host "Logica para reapuntar IDs Lookup erroneos del campo 'Seccion' en la biblioteca '$NombreDeBiblioteca' finalizada con exito" -ForegroundColor Green
    Write-Log -Nivel "INFO" -Mensaje "Logica para reapuntar IDs Lookup erroneos del campo 'Seccion' en la biblioteca '$NombreDeBiblioteca' finalizada con exito"
    # ==========================================
    # FIN DE LOGICA PARA REAPUNTAR IDs LOOKUP PARA EL CAMPO FdsSeccion EN LA BIBLIOTECA
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
