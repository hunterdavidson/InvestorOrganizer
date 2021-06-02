####Enter excel file name here
$file = "Investor Excel File"

####Enter starting row (rowA) and ending row (rowB)
$rowA = 0
$rowB = 0

####Enter starting collumn (collumnA) and ending collumn (collumnB)
$collumnA = 0
$collumnB = 0


$workingDir = Get-Location
$excelObj = New-Object -ComObject Excel.Application
$excelObj.Visible = $false
$workBook = $excelObj.Workbooks.Open($workingDir+$file)
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

Write-Host "Folder Creation Completed"
cmd /c 'pause'