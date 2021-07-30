$InvestorFoldersPATH = "Z:\Shared\Rainier Medical\Accounting-Intern File 2021\NEMI4 Folders\NEMI4-Investor Files"

$InvestorFilesPATH = "Z:\Shared\Rainier Medical\Accounting-Intern File 2021\NEMI4 Folders\Investor Files"

$InvestorNumberList = New-Object Collections.Generic.List[String]

foreach($folder in Get-ChildItem $InvestorFoldersPATH) {

   $folderCount = Get-ChildItem -path $folder.FullName | Measure-Object

    if ($folderCount.Count -le 3) {

        $FolderName = $folder.Name
        
        if ($FolderName -match "NEMI4 (?<investorNumber>.*) - ") {
            $InvestorNumberList.Add("NEMI4-" + $matches['investorNumber'])
        }

    }
}

$InvestorExcelPATH = "Z:\Shared\Rainier Medical\Accounting-Intern File 2021\NEMI4 Folders\NEMI4-Doc Index by Investor Name-07.29.2021-A.csv"

$ExcelRows = New-Object Collections.Generic.List[Int]

$CurrentDocumentIDs = New-Object Collections.Generic.List[String]

$DocumentUIDs = New-Object Collections.Generic.List[String]

$FileNames = New-Object Collections.Generic.List[String]

$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open($InvestorExcelPATH)
$Excel.visible = $false

foreach ($number in $InvestorNumberList) {
    $numberStore = $number.Split("-", 2)

    $Target = $Workbook.Sheets.Item(1).UsedRange.Find($number)
    $number
    $First = $Target

    $name = ""

    Do
    {
        $ExcelRows.Add($Target.Row)
        $Target = $Workbook.Sheets.Item(1).UsedRange.FindNext($Target)
    }

    While ($Target -ne $NULL -and $Target.AddressLocal() -ne $First.AddressLocal())

    foreach($row in $ExcelRows) {

       $CurrentDocumentIDs.Add($Workbook.Sheets.Item(1).Columns.Item(1).Rows.Item($row).Text)
       $name = $Workbook.Sheets.Item(1).Columns.Item(8).Rows.Item($row).Text
       $DocumentUIDs.Add($Workbook.Sheets.Item(1).Columns.Item(14).Rows.Item($row).Text)

    }

    $folderPath = $InvestorFoldersPATH+"\"+$numberStore+" - "+$name

    for($counter=0; $counter -lt $CurrentDocumentIDs.Count; $counter++){

        $currentDocumentID = $CurrentDocumentIDs[$counter]

        Get-ChildItem -Path $InvestorFilesPATH -Recurse -Filter $currentDocumentID*.pdf | 
            Foreach-Object {

                Move-Item $_.FullName -Destination $folderPath
                
                Get-ChildItem -Path $folderPath -Recurse -Filter $currentDocumentID*.pdf |  Foreach-Object {
                    if ((Rename-Item -Path $_.FullName -NewName $currentDocumentUID` $numberStore` -` $name`.pdf -ErrorAction Ignore) -eq $null) {
                        Remove-Item -Path $_.FullName -ErrorAction Ignore
                    }
                }
	        }
    }

    $ExcelRows = New-Object Collections.Generic.List[Int]
    $CurrentDocumentIDs = New-Object Collections.Generic.List[String]
    $DocumentUIDs = New-Object Collections.Generic.List[String]
}

$Workbook.Close()
$Excel.Quit()