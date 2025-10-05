$SiteURL = "https://circo.sharepoint.com/sites/SHP-YPF-DEV/"

# Conectar a SharePoint Online
Connect-PnPOnline -Url $SiteURL -UseWebLogin

# Obtener todas las listas del sitio
$lists = Get-PnPList

foreach ($list in $lists) {
    if ($list.Title -like "OBIRAS*") {
        Write-Host "Procesando lista: $($list.Title)" -ForegroundColor Cyan

        # Traer todos los campos de la lista
        $fields = Get-PnPField -List $list

        foreach ($field in $fields) {
            if ($field.InternalName -like "Bloque*") {
                Write-Host "  Eliminando campo: $($field.InternalName)" -ForegroundColor Red
                Remove-PnPField -List $list -Identity $field.InternalName -Force
            }
        }
    }
}

Disconnect-PnPOnline
Write-Host "Finalizado." -ForegroundColor Green
