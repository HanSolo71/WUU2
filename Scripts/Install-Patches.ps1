$updateSession = New-Object -ComObject 'Microsoft.Update.Session'
$updateSearcher = $updateSession.CreateupdateSearcher()
$searchResult = $updateSearcher.Search("IsInstalled=0 and IsHidden=0")

if($searchResult.updates.count -eq 0){return 0}

$updatesToInstall = New-Object -ComObject "Microsoft.Update.UpdateColl"

ForEach($update in $searchResult.Updates){
    if($update.InstallationBehavior.CanRequestUserInput -eq $true){continue}
    if($update.IsDownloaded -eq $false){continue}
    if($update.EulaAccepted -eq $false){$update.AcceptEula()}
    $updatesToInstall.Add($update) | Out-Null
}

$installer = $updateSession.CreateUpdateInstaller()
$installer.Updates = $updatesToInstall
$installationResult = $installer.Install()

$numErrors = 0
0..($updatesToInstall.Count - 1) | % {
    if($installationResult.GetUpdateResult($_).ResultCode -ge 4){$numErrors++}
}
return $numErrors