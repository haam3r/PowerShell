#Requires -Version 2.0

<#
.Synopsis
   Batch adding mail contacts to Microsoft Outlook based on list from CSV file.
.DESCRIPTION
   Takes input from CSV file. Adds everybody in the CSV list as a contact.
   The CSV file is expected to have a structure like this: "Email1Address;FullName", including this header.
   How to run: "powershell.exe -executionpolicy bypass -command ". .\Set-OutlookProfile.ps1; Set-OutlookProfile -PRF .\peeter_nuianken.prf""
.PARAMETER PRF
    PRF file for profile setup. See "https://technet.microsoft.com/en-us/library/cc179062(v=office.14).aspx" on how to create.
.PARAMETER OutlookPath
    Path to folder containing 'OUTLOOK.EXE', default's to 64bit Office 2010 aka 'C:\Program Files\Microsoft Office\Office14'
.EXAMPLE
   Set-OutlookProfile -PRF C:\Users\admin\Desktop\account.PRF

   Use the selected PRF file to set up an Outlook profile.
.EXAMPLE
    Set-OutlookProfile -PRF C:\Users\admin\Desktop\account.PRF -OutlookPath "C:\Program Files\Microsoft Office\Office14"

    Use the selected PRF file to set up an Outlook profile. Use the specified path to OUTLOOK.EXE aka specify Office version.
.NOTES
   Author: Andres Elliku
   Version: 1.0
#>
function Set-OutlookProfile {
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   HelpMessage="PRF file for profile")]
        [ValidateScript({Test-Path $_ -PathType 'Leaf'})]
        [Alias("P")]
        [string]$PRF,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1,
                   HelpMessage="Path to folder containing OUTLOOK.EXE")]
        [ValidateScript({Test-Path $_ -PathType 'Container'})]
        [Alias("P","Path")]
        [string]$OutlookPath = "C:\Program Files\Microsoft Office\Office14"
    )
    cd $OutlookPath
    Start-Process -FilePath ".\OUTLOOK.EXE" -Argumentlist "/importprf $PRF"
    # Outlook startup is slow. Give it time to actually do the import.
    sleep 2
    Stop-Process -Name "OUTLOOK*"
}