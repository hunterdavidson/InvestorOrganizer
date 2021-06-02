$workingDir = Get-Location

####Investor Fund Name
$InvestorFundName = "New Era Medical Investment Fund IV"


foreach($folder in Get-ChildItem $workingDir\InvestorFolders) {
    $currentInvestorName = $folder.ToString().Split("-", 2)[1].Trim()

    Get-ChildItem –Path $workingDir\InvestorFiles -Recurse -Filter *-$InvestorFundName-$currentInvestorName.pdf | Foreach-Object {Copy-Item $_.FullName -Destination $workingDir\InvestorFolders\$folder}
}

Write-Host "Sorting Completed"
cmd /c 'pause'