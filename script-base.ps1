try {
    Write-Host "Conectando a SharePoint Online..." -f Magenta
    $SiteURL = "https://circo.sharepoint.com/sites/SPXTest/"
    $WarningPreference = "SilentlyContinue"
    Connect-PnPOnline -Url $SiteURL -UseWebLogin

    Write-Host "Accion a ejecutar..." -f Cyan



    Write-Host "Script ejecutado satisfactoriamente." -f Green
} Catch {
    Write-Host -f Red "¡Error!" $_.Exception.Message
} Finally {
    Disconnect-PnPOnline
    Write-Host "Conexión cerrada." -f Magenta
}

# Magenta   -> recursos
# Cyan      -> acciones
# Green     -> acciones completadas con exito
# Red       -> errores
# Yellow    -> warnings
