#Requires -Version 2.0

<#
.Synopsis
   Get example CSV file, with all possible Outlook contact fields.
.DESCRIPTION
   Get example CSV file, with all possible Outlook contact fields.
.PARAMETER FilePath
    Path where to output CSV
.EXAMPLE
    Get-ContactProperties -FilePath "C:\example.CSV"
.NOTES
   Author: Andres Elliku
   Version: 1.0
   Source: https://itmicah.wordpress.com/2013/11/14/add-contacts-to-outlook-using-powershell-and-a-csv-file/
#>
function Get-ContactProperties {
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   HelpMessage="Full path where to output CSV.")]
        [Alias("P","Path")]
        [string]$FilePath = "C:\Users\admin\Desktop\example.csv"
    )
    $outlook = new-object -com Outlook.Application -ea 1
    $contacts = $outlook.session.GetDefaultFolder(10)
    $newcontact = $contacts.Items.Add()
    $Props = $newcontact | gm -MemberType property | ?{$_.definition -like 'string*{set}*'}
    $newcontact.Delete()
    $properties = $Props | ForEach-Object {$_.Name}
    $properties | select $properties | Export-Csv $FilePath -UseCulture -NoTypeInformation
}