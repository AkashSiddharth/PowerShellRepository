function Load-Module {
    param (
        [Parameter(Mandatory=$true)][string]$ModuleName
    )

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object {$_.Name -eq $ModuleName}) {
        write-host "Module $ModuleName is already imported."
    }
    else {
        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $ModuleName}) {
            Import-Module $ModuleName -Verbose
        }
        else {
            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $moduleName | Where-Object {$_.Name -eq $ModuleName}) {
                Install-Module -Name $ModuleName -Force -Verbose -Scope CurrentUser
                Import-Module $ModuleName -Verbose
            }
            else {
                # If module is not imported, not available and not in online gallery then abort
                write-host "Module $ModuleName not imported, not available and not in online gallery, exiting."
                EXIT 1
            }
        }
    }
}

$Excel = New-Object -ComObject "Excel.Application"

$Excel.Visible = $true
$Excel.DisplayAlerts = $false

$FileName = "E:\TemporaryFiles\output.xlsx"
# Open the Excel File
$WB = $Excel.Workbooks.Open($FileName)


$WS = $WB.Worksheets.Sheets | Where {$_.Name -eq "Test"}
#$WS.Name = "TestNew"

$WB.Save()
# Cleaning up
# $WB.Close()
# $Excel.Quit()

# Stopping Process
#Get-Process excel | Stop-Process -Force

# Garbage Collection
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)