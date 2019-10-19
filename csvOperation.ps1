param (
    [Parameter(Mandatory=$true)][String]$Target,
    [Parameter(Mandatory=$true)][String]$InputColumn
)

# Get all CSV files
Get-ChildItem -Path $Target/* -Include *.csv | ForEach-Object {
    $CsvFile = Import-Csv $_.FullName

    # Define an empty array
    $DataArray = $CsvFile.Country

    # Get all the values of the requested column
    # ForEach-Object {
    #     Write-Output "$CsvFile.$Column"
    #     $DataArray += $($CsvFile.$Column)
    # }
    # Write-Output $DataArray

    $joinedVal = [System.String]::Join(" ", $DataArray)
    # Write-Debug "Get-CSVStringData:: Data: $joinedVal"

    $ReturnHash = @{
        Count = $DataArray.count
        Value = $joinedVal
    }

    # Testing the structure
    $ReturnHash.keys | ForEach-Object {
        Write-Output "Key: $_"
        Write-Output "Value: $($ReturnHash[$_])"
        Write-Output "-------------------------------------"
    }
}