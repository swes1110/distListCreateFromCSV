<#
.SYNOPSIS
    Create distrubition list from given csv of addresses

.DESCRIPTION


.PARAM csvPath
    Path to CSV file containing Names/Addresses to add

.PARAM listName
    Name of the dist list that you would like to create

.PARAM generateCsv
    Output csv file to be filled in

.NOTES
    Filename: CreateDistList.ps1
    Version: 0.1
    Author: Shawn Reynolds
    Organization: The American Board of Family Medicine, Inc.
    Blog: swes1110.blogspot.com
    Twitter: @swes1110

    Version history:

    0.1     -   Script created
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$csvPath,
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$listName,
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty]
    [switch]$exportCSV
)

if ($exportCSV) {
    # Code to generate CSV file
}

# Import exchange powershell module (must be installed on local machine)
try {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
catch [System.Exception]{
    Write-Warning -Message "Failed to add Exchange Management snap-in. Is it installed on this computer?"
    Write-Warning -Message $_.Exception.Message
    break
}

# Import CSV from path given as argument
function csvImport {
    param (
        $csvPath
    )
    try {
        $importedCSV = Import-Csv -Path $csvPath
        return $importedCSV
    }
    catch [System.Exception]{
        Write-Warning -Message "Failed to import csv from $csvPath"
        Write-Error -Message $_.Exception.Message
        breakPr
    }
}

function checkContact {
    param (
        $eMailAddress
    )
    try {
        $output = Get-MailContact -Filter "PrimarySmtpAddress -eq '$eMailAddress'"
        Write-Output $output
    }
    catch [System.Exception]{
        Write-Warning -Message "Error getting contact $eMailAddress"
        Write-Error -Message $_.Exception.Message
    }
}

$test = csvImport $csvPath
foreach ($email in $test.eMailAddress) {
    checkContact $email    
}

git config --global user.email "you@example.com"
git config --global user.name "Your Name"