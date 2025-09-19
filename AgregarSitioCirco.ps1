[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteName
)

try {
    $WarningPreference = "SilentlyContinue"
    Connect-PnPOnline -Url "https://circo-admin.sharepoint.com" -UseWebLogin
    Write-Host "Conexión abierta.`n" -ForegroundColor Magenta

    Write-Host "Agregando sitio a PnP SiteCollectionAppCatalog..." -ForegroundColor Cyan
    Add-PnPSiteCollectionAppCatalog -Site "https://circo.sharepoint.com/sites/$SiteName/"

    Write-Host "`nScript ejecutado satisfactoriamente.`n" -ForegroundColor Green
}
catch {
    Write-Host -ForegroundColor Red "Hubo un error: $($_.Exception.Message)"
}
finally {
    Disconnect-PnPOnline
    Write-Host "Conexión cerrada." -ForegroundColor Magenta
}
