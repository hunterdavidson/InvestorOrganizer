####Enter starting row (rowA) and ending row (rowB)
$rowA = 0
$rowB = 0

####Enter starting collumn (collumnA) and ending collumn (collumnB)
$collumnA = 0
$collumnB = 0

function Ask-Continue {
     if ((Read-Host "Press enter to continue or type 'exit' to exit") -eq "exit") {
          End-Program
     }
}

$workingDir = Get-Location
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

function End-Program {$excel.Quit()
     $excel.Quit()
     [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

     Write-Host "Folder Creation Completed"
     cmd /c 'pause'
}

try { $workBook = $excel.Workbooks.Open($workingDir.ToString()+"\InvestorNames\*.xlsx") }
catch { 
     Write-Host There was an error opening excel file at $workingDir.ToString()+"\InvestorNames\*.xlsx"
     End-Program
}

$workSheet = $workBook.sheets.Item(1)

$folderName = ""

$continue = $false


for($i=$rowA;$i-le $rowB;$i++){

     for($a=$collumnA;$a-le $collumnB;$a++){
        $folderName += $workSheet.Columns.Item($a).Rows.Item($i).Text.Trim()

        if($folderName -eq "") {
            Write-Host Missing data in row[$i] -> Folder not created
            Ask-Continue
            $continue = $true
	        break
        }

        if($a -lt $collumnB){
            $folderName += " "
        }
     }

     if($continue -eq $true) {
            $continue = $false
            continue
        }
    
     
     try { 
        $NewItem = New-Item -path $workingDir\InvestorFolders\$folderName -ItemType "directory" -ErrorAction Ignore
        if ($null -eq $NewItem) {
            throw
            }
        }
     catch {
	    Write-Host Folder [$folderName] has already been created or could not be created because of an error
        $folderName = ""
        continue
     }

     $folderName = ""
}

End-Program
