$workingDir = Get-Location

foreach($folder in Get-ChildItem $workingDir\InvestorFolders) {
   
    if($folder.ToString() -match "placeholder") { continue }
    $currentInvestorName = $folder.ToString().Split("-", 2)[1].Trim()

    Get-ChildItem -Path $workingDir\InvestorFiles -Recurse -Filter *-$currentInvestorName.pdf | 
        Foreach-Object {
            if($_.FullName.ToString() -match "1st Amendment") { continue }
            $_.FullName.ToString()
            if($_.FullName.ToString() -match "Fully Executed") {Copy-Item $_.FullName -Destination $workingDir\InvestorFolders\$folder}
            }
}

Write-Host "Sorting Completed"
cmd /c 'pause'
