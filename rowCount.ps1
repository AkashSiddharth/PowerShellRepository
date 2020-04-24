param (
    [Parameter(Mandatory=$true)][String]$Target,
    [Parameter(Mandatory=$true)][String]$OutputFile
)

# Define an empty HashTable
$ReturnHash = @{}

# Get all CSV files in the directory
Get-ChildItem -Path $Target/* -Include *.csv | ForEach-Object {
    # Get the row Count and Data
    $CSVFile = $_.FullName

    Write-Debug "CSV File in this iteration: $CSVFile"

    # Count the lines
    $Lines = (Import-Csv $CSVFile).count

    # Update the HashTable
    $ReturnHash[$CSVFile] = $Lines
}

# Write to file
Out-File -FilePath $OutputFile -InputObject $ReturnHash
