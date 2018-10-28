$updateSession = New-Object -ComObject 'Microsoft.Update.Session'
$updateSearcher = $updateSession.CreateupdateSearcher()
$searchResult = $updateSearcher.Search("IsInstalled=0 and IsHidden=0")
$updatesToDownload = New-Object -ComObject "Microsoft.Update.UpdateColl"
ForEach($update in $searchResult.Updates){
    if($update.IsDownloaded -eq $false){
        $updatesToDownload.Add($update) | Out-Null
    }
}
if($updatesToDownload.Count -gt 0){
    $downloader = $updateSession.CreateUpdateDownloader()
    $downloader.Updates = $updatesToDownload
    $downloadResult = $downloader.Download()

    $numDownloaded = 0
    0..($updatesToDownload.Count - 1) | % {
        $result = $downloadResult.GetUpdateResult($_).ResultCode
        if($result -eq 2 -or $result -eq 3){$numDownloaded++}
    }
    return $numDownloaded
}
else{return 0}