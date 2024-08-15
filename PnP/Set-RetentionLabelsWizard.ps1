##Adding Retention Labels to Document Libraries using PowerShell

$TenantName = Read-Host "Enter the tenant name (e.g. if the sharepoint url is contoso.sharepoint.com, enter 'contoso')"

Connect-pnponline -url "https://$tenantname.sharepoint.com" -Interactive

[array]$Sites = (Get-PnPTenantSite | Out-GridView -PassThru)

Write-Host "You have chosen to process $($Sites.count) site(s)"

Write-Host "Select the Retention Label you would like to apply"

$Label = Get-PnPLabel | Select tagname,notes | Out-GridView -OutputMode Single -Title "Choose a retention label to apply"

foreach($site in $sites){

    Connect-PnPOnline -url $site.url -Interactive

    [array]$Libraries = get-pnplist | ?{$_.basetype -eq "DocumentLibrary" -and (!$_.hidden)} | select  title,id| Out-GridView -OutputMode Single -Title "Select one or more libraries to apply the label $($label.tagname) to"

    foreach($library in $libraries){

    Set-PnPLabel -List $library.title -Label $label.tagname

    }

}