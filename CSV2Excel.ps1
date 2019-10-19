param (
    [Parameter(Mandatory=$true)][String]$Target,
    [Parameter(Mandatory=$true)][String]$OutputExcelFile,
    [Parameter(Mandatory=$true)][String]$OutputWorksheet,
    [Parameter(Mandatory=$true)][Int]$CountColumn,
    [Parameter(Mandatory=$true)][Int]$DataColumn
)

# The Function returns the number of rows and the string representation of all the
# data from the specified column
# RETURNS :: Hash {No. of rows : String Rows Value}
function Get-CSVStringData {
    param (
        [Parameter(Mandatory=$true)][String]$FileName
    )
    # Write-Debug "Get-CSVStringData:: File: $FileName"

    begin{
        # Open CSV files
        $CsvFile = Import-Csv $FileName

        # Define an empty array
        $DataArray = @()

        # Define an emprty HashTable
        $ReturnHash = @{}
    }
    process{
        # Get the column values into the array
        $DataArray = $CsvFile.Country

        # Get the requisites into the Hash
        $ReturnHash["Count"] = $DataArray.count
        $ReturnHash["Value"] = [System.String]::Join(" ", $DataArray)
    }
    end{
        Return $ReturnHash
    }
}

# Does not use the ImportExcel module, Uses Excel COM module
# Opens Excel Application in Visible mode and write data to it
function Write-ToExcel {
    param (
        [Parameter(Mandatory=$true)][string]$FileName,
        [Parameter(Mandatory=$true)][string]$SheetName,
        [Parameter(Mandatory=$true)][Int]$CountColumn,
        [Parameter(Mandatory=$true)][Int]$DataColumn,
        [Parameter(Mandatory=$true)][Hashtable]$ToWrite
    )

    # Write-Debug "Write-ToExcel :: Param :: FileName -> $FileName"
    # Write-Debug "Write-ToExcel :: Param :: SheetName -> $SheetName"
    # Write-Debug "Write-ToExcel :: Param :: CountColumn -> $CountColumn"
    # Write-Debug "Write-ToExcel :: Param :: DataColumn -> $DataColumn"
    # Write-Debug "Write-ToExcel :: Param :: ToWrite -> "
    # $ToWrite.keys | ForEach-Object {
    #     Write-Debug "Key: $_"
    #     Write-Debug "Value: $($ToWrite[$_])"
    #     Write-Debug "-------------------------------------"
    # }
    begin {
        $Excel = New-Object -ComObject "Excel.Application"

        $Excel.Visible = $true
        $Excel.DisplayAlerts = $false

        # Open the Excel File
        $WB = $Excel.Workbooks.Open($FileName)
    }
    process{
        # Activate the WorkSheet
        $WS = $WB.Sheets | Where {$_.Name -eq $SheetName}
        $Cells = $WS.Cells

        #Insert the data into the column
        ### Get the Last cell
        $WSRange = $WS.UsedRange 
        $EndRow = $WSRange.SpecialCells(11).row

        # Enter the Count value
        $EndRow++
        $Cells.item($EndRow, $CountColumn) = $ToWrite['Count']

        # Enter the Data value
        $Cells.item($EndRow, $DataColumn) = $ToWrite['Value']
    }
    end{
        # Save the file
        $WB.Save()

        # Cleaning up
        $WB.Close()
        $Excel.Quit()
    }
}

# Get all CSV files
Get-ChildItem -Path $Target/* -Include *.csv | ForEach-Object {
    # Get the row Count and Data
    $CSVFile = $_.FullName

    Write-Debug "CSV File in this iteration: $CSVFile"
    $Outcome = Get-CSVStringData -FileName $CSVFile

    # Testing the structure
    # $Outcome.keys | ForEach-Object {
    #     Write-Debug "Key: $_"
    #     Write-Debug "Value: $($Outcome[$_])"
    #     Write-Debug "-------------------------------------"
    # }

    # Write to Excel
    Write-ToExcel -FileName $OutputExcelFile -SheetName $OutputWorksheet -CountColumn $CountColumn -DataColumn $DataColumn -ToWrite $Outcome
}
