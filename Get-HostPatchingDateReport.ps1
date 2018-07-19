param(
    [Parameter(Mandatory=$true)][String]$vCenter
    )

# -----------------------
# Define Global Variables
# -----------------------
$Global:Folder = $env:USERPROFILE+"\Documents\HostPatchingReports" 

#*************************************************
# Check for Folder Structure if not present create
#*************************************************
Function Verify-Folders {
    [CmdletBinding()]
    Param()
    "Building Local folder structure" 
    If (!(Test-Path $Global:Folder)) {
        New-Item $Global:Folder -type Directory
  <#      New-Item "$Global:WorkFolder\Annotations" -type Directory
        New-Item "$Global:WorkFolder\IPConfig" -type Directory
        New-Item "$Global:WorkFolder\ShareInfo" -type Directory
        New-Item "$Global:WorkFolder\VMInfo" -type Directory
        New-Item "$Global:WorkFolder\PrinterInfo" -type Directory
    #>
        }
    "Folder Structure built" 
}
#***************************
# EndFunction Verify-Folders
#***************************

#*******************
# Connect to vCenter
#*******************
Function Connect-VC {
    [CmdletBinding()]
    Param()
    "Connecting to $Global:VCName"
    Connect-VIServer $Global:VCName -Credential $Global:Creds -WarningAction SilentlyContinue
}
#***********************
# EndFunction Connect-VC
#***********************

#*******************
# Disconnect vCenter
#*******************
Function Disconnect-VC {
    [CmdletBinding()]
    Param()
    "Disconnecting $Global:VCName"
    Disconnect-VIServer -Server $Global:VCName -Confirm:$false
}
#**************************
# EndFunction Disconnect-VC
#**************************

#**************************
# Function Check-PowerCLI10 
#**************************
Function Check-PowerCLI10 {
    [CmdletBinding()]
    Param()
    #Check for Prereqs for the script
    #This includes, PowerCLI 10, plink, and pscp

    #Check for PowerCLI 10
    $powercli = Get-Module -ListAvailable VMware.PowerCLI
    if (!($powercli.version.Major -eq "10")) {
        Throw "VMware PowerCLI 10 is not installed on your system!!!"
    }
    Else {
        Write-Host "PowerCLI 10 is Installed" -ForegroundColor Green
    } 
}
#*****************************
# EndFunction Check-PowerCLI10
#*****************************

#**************************
# Function Convert-To-Excel
#**************************
Function Convert-To-Excel {
    [CmdletBinding()]
    Param()
   "Converting HostList from $Global:VCname to Excel"
    $workingdir = $Global:Folder+ "\*.csv"
    $csv = dir -path $workingdir

    foreach($inputCSV in $csv){
        $outputXLSX = $inputCSV.DirectoryName + "\" + $inputCSV.Basename + ".xlsx"
        ### Create a new Excel Workbook with one empty sheet
        $excel = New-Object -ComObject excel.application 
        $excel.DisplayAlerts = $False
        $workbook = $excel.Workbooks.Add(1)
        $worksheet = $workbook.worksheets.Item(1)

        ### Build the QueryTables.Add command
        ### QueryTables does the same as when clicking "Data » From Text" in Excel
        $TxtConnector = ("TEXT;" + $inputCSV)
        $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
        $query = $worksheet.QueryTables.item($Connector.name)


        ### Set the delimiter (, or ;) according to your regional settings
        ### $Excel.Application.International(3) = ,
        ### $Excel.Application.International(5) = ;
        $query.TextFileOtherDelimiter = $Excel.Application.International(5)

        ### Set the format to delimited and text for every column
        ### A trick to create an array of 2s is used with the preceding comma
        $query.TextFileParseType  = 1
        $query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        ### Execute & delete the import query
        $query.Refresh()
        $query.Delete()

        ### Get Size of Worksheet
        $objRange = $worksheet.UsedRange.Cells 
        $xRow = $objRange.SpecialCells(11).ow
        $xCol = $objRange.SpecialCells(11).column

        ### Format First Row
        $RangeToFormat = $worksheet.Range("1:1")
        $RangeToFormat.Style = 'Accent1'

        ### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
        $Workbook.SaveAs($outputXLSX,51)
        $excel.Quit()
    }
    ## To exclude an item, use the '-exclude' parameter (wildcards if needed)
    remove-item -path $workingdir 

}
#*****************************
# EndFunction Convert-To-Excel
#*****************************

#******************************
# Function Get-HostPatchingDate
#******************************
Function Get-HostPatchingDate {
    [CmdletBinding()]
    Param()
    "Getting Listing of last Patched dates for Hosts in $Global:VCName"
    $PatchInfo = @()
    $Count = 1
    $VMHosts = Get-VMHost

    ForEach ($vmhost in $VMHosts){
        Write-Progress -Id 0 -Activity 'Generating Patch Date Details ' -Status "Processing $($count) of $($VMHosts.count)" -CurrentOperation $_.Name -PercentComplete (($count/$VMHosts.count) * 100)
        $into = New-Object PSObject
        $LastPatchDate = [datetime]((Get-ESXCli -VMHost $vmhost).software.vib.list() |Select-Object -Property installdate -ExpandProperty installdate|Sort-Object -Descending)[0]
        
        Add-Member -InputObject $into -MemberType NoteProperty -Name Host -Value $vmhost.Name
        Add-Member -InputObject $into -MemberType NoteProperty -Name Version -Value $vmhost.version
        Add-Member -InputObject $into -MemberType NoteProperty -Name Build -Value $vmhost.build
        Add-Member -InputObject $into -MemberType NoteProperty -Name LastPatchDate -Value $LastPatchDate
        $PatchInfo += $into
        $into = $null
        $Count++
    }
    Write-Progress -Id 0 -Activity 'Generating Patch Date Details ' -Completed
    $PatchInfo | Export-CSV -Path $Global:Folder\$Global:VCname-PatchingDates.csv -NoTypeInformation

}
#*********************************
# EndFunction Get-HostPatchingDate
#*********************************


#***************
# Execute Script
#***************
CLS
"=========================================================="
#Verify all require software is installed
"Checking for required Software on your system"
"=========================================================="
Check-PowerCLI10
"=========================================================="
Verify-Folders
#$ErrorActionPreference="SilentlyContinue"

"=========================================================="
" "
Write-Host "Get CIHS credentials" -ForegroundColor Yellow
$Global:Creds = Get-Credential -Credential $null

#Get-VCenter
$Global:VCName = $vCenter
Connect-VC
"----------------------------------------------------------"
Get-HostPatchingDate
Disconnect-VC
Convert-To-Excel
"Open Explorer to $Global:Folder"
Invoke-Item $Global:Folder