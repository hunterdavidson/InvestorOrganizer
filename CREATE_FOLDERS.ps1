####Enter excel file name here
$file = "Investor Excel File"

####Enter starting row (rowA) and ending row (rowB)
$rowA = 0
$rowB = 0

####Enter starting collumn (collumnA) and ending collumn (collumnB)
$collumnA = 0
$collumnB = 0


$workingDir = Get-Location
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workBook = $excel.Workbooks.Open($workingDir+"\InvestorNames\"+$file)
$workSheet = $workBook.sheets.Item(1)

$folderName = ""


for($i=$rowA;$i-le $rowB;$i++){

     for($a=$collumnA;$a-le $collumnB;$a++){
        $folderName += $workSheet.Columns.Item($a).Rows.Item($i).Text.Trim()

        if($folderName -eq "") {
            Write-Host Missing data at collumn [$a], row[$i]
        }

        if($a -lt $collumnB){
            $folderName += " "
        }
     }

     if($folderName -eq "") {
            Write-Host Missing data at row[$i] -> Folder not created
            continue
        }
    
     
     New-Item -path InvestorFolders\$folderName -ItemType "directory"
     $folderName = ""
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel

Write-Host "Folder Creation Completed"
cmd /c 'pause'