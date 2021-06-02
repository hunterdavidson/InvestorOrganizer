$workingDir = Get-Location


foreach($folder in Get-ChildItem $workingDir\InvestorFolders) {

    Get-ChildItem –Path $folder.FullName -Recurse -Filter *.pdf | Foreach-Object {Start-Process -FilePath $_.FullName –Verb Print -PassThru | %{sleep 4;$_} | kill}
}

Write-Host "Printing Completed"
cmd /c 'pause'