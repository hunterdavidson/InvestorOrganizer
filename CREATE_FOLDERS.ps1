####Enter starting row (rowA) and ending row (rowB)
$rowA = 192
$rowB = 206

####Enter starting collumn (collumnA) and ending collumn (collumnB)
$collumnA = 1
$collumnB = 3


$workingDir = Get-Location
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workBook = $excel.Workbooks.Open($workingDir.ToString()+"\InvestorNames\*.xlsx")
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
    
     
     New-Item -path $workingDir\InvestorFolders\$folderName -ItemType "directory"
     $folderName = ""
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel

Write-Host "Folder Creation Completed"
cmd /c 'pause'
