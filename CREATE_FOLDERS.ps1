####Enter starting row (rowA) and ending row (rowB)
$rowA = 1
$rowB = 16

####Enter starting collumn (collumnA) and ending collumn (collumnB)
$collumnA = 1
$collumnB = 3

function End-Program {
     $excel.Quit()
     [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
     Remove-Variable $excel

     Write-Host "Folder Creation Completed"
     cmd /c 'pause'
}

function Ask-Continue {
     if (Read-Host "Press enter to continue or type 'exit' to exit." -eq "exit") {
          End-Program
     }
}

$workingDir = Get-Location
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

try { $workBook = $excel.Workbooks.Open($workingDir.ToString()+"\InvestorNames\*.xlsx") }
catch { 
     Write-Host There was an error opening excel file at $workingDir.ToString()+"\InvestorNames\*.xlsx"
     End-Program
}

$workSheet = $workBook.sheets.Item(1)

$folderName = ""


for($i=$rowA;$i-le $rowB;$i++){

     for($a=$collumnA;$a-le $collumnB;$a++){
        $folderName += $workSheet.Columns.Item($a).Rows.Item($i).Text.Trim()

        if($folderName -eq "") {
            Write-Host Missing data at collumn [$a], row[$i]
	    $folderName = ""
            Ask-Continue
	    continue
        }

        if($a -lt $collumnB){
            $folderName += " "
        }
     }

     if($folderName -eq "") {
            Write-Host Missing data at row[$i] -> Folder not created
            Ask-Continue
	    $folderName = ""
            continue
        }
    
     
     try { New-Item -path $workingDir\InvestorFolders\$folderName -ItemType "directory" }
     catch {
	    Write-Host Folder has already been created or could not be created because of an error
        continue
     }

     $folderName = ""
}

End-Program
