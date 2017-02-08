#Requires -Version 2.0
#Requires -RunAsAdministrator

function Get-ContactProperties {
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

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   HelpMessage="Full path where to output CSV.")]
        [ValidateScript({ Test-Path -Path (Split-Path -Parent $_ -OutVariable Parent) -PathType Container })]
        [Alias("P","Path")]
        [string]$FilePath = "$env:USERPROFILE\Desktop\example.csv"
    )

    try {
        Write-Verbose -Message "Trying to create a com object to communicate with Outlook"
        $Outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
    }
    catch {
        $Error[0]
    }

    Write-Verbose -Message "Querying the default folder"
    $Contacts = $Outlook.session.GetDefaultFolder(10)
    
    Write-Verbose -Message "Adding temporary contact"
    $NewContact = $Contacts.Items.Add()
    $Members = $NewContact | Get-Member -MemberType Property | Where-Object { $_.definition -like 'string*{set}*'}
    
    Write-Verbose -Message "Delete temporary contact"
    $NewContact.Delete()
    
    Write-Verbose -Message "Enumerating all possible properties and outputting them to file"
    $Properties = $Members | ForEach-Object {$_.Name}
    try {
        $Properties | Select-Object $Properties | Export-Csv $FilePath -UseCulture -NoTypeInformation -Encoding UTF8
    }
    catch {
        $Error[0]
    }
}


function Set-OutlookContact {
<#
.Synopsis
   Batch adding mail contacts to Microsoft Outlook based on list from CSV file.
.DESCRIPTION
   Takes input from CSV file. Adds everybody in the CSV list as a contact.
   The CSV file is expected to have a structure like this: "Email1Address;FullName", including this header.
   How to run: "powershell.exe -executionpolicy bypass -command ". .\Set-OutlookContact.ps1; Set-OutlookContact -CSV .\test.csv""
.PARAMETER CSV
    Source csv file from wich to extract user list.The CSV file is expected to have a structure like this: "Email1Address;FullName", including this header.
    Get-ContactProperties will give you an example csv file with all possible header values.
.EXAMPLE
   Set-OutlookContact -CSV .\list.csv

   Take contacts from csv file "list.csv" and create them in Outlook contacts list.
.NOTES
   Author: Andres Elliku
   Version: 1.0
   Source: https://itmicah.wordpress.com/2013/11/14/add-contacts-to-outlook-using-powershell-and-a-csv-file/
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   HelpMessage="Source csv file from wich to extract user list")]
        [ValidateScript({Test-Path $_ -PathType 'Leaf'})]
        [Alias("C","Path")]
        [string]$CSV
    )

    Begin {

        Write-Verbose -Message "Importing from $CSV"
        $Contacts = Import-Csv -Path $CSV -Delimiter ";"

        try {
            Write-Verbose -Message "Trying to create a com object to communicate with Outlook"
            $Outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
        }
        catch {
            $Error[0]
        }

        Write-Verbose "Setting up a contacts session"
        $DefaultFolder = $Outlook.session.GetDefaultFolder(10)
    }

    Process {
        
        Write-Verbose "Add contacts with their properties"
        foreach ($Contact in $Contacts) {
            
            $NewContact = $DefaultFolder.Items.Add()
            foreach ($Property in $Contact.PSobject.Properties) {
                $NewContact.$($Property.Name) = $Property.Value
            }
            $NewContact.Save()
            $AddedContact += $($NewContact.Email1Address)
            Write-Verbose $AddedContact
        }
    }

    End {
        $CountTotal = ($Contacts | Measure-Object).Count
        Write-Verbose -Message "Total number of contacts: $count"
        $CountAdded = ($AddedContact | Measure-Object).Count
        Write-Verbose -Message "Number of added contacts: $CountAdded"

        Write-Verbose "Close Outlook connection and remove com object"
        $Outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)
    }

}


function Set-OutlookProfile {
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

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   HelpMessage="PRF file for profile")]
        [ValidateScript({Test-Path $_ -PathType 'Leaf'})]
        [Alias("P","Path")]
        [string]$PRF,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1,
                   HelpMessage="Path to folder containing OUTLOOK.EXE")]
        [ValidateScript({Test-Path $_ -PathType 'Container'})]
        [string]$OutlookPath = "C:\Program Files\Microsoft Office\Office14"
    )

    Set-Location -Path $OutlookPath
    Write-Verbose -Message "Starting Outlook and importing profile"
    Try {
        Start-Process -FilePath ".\OUTLOOK.EXE" -Argumentlist "/importprf $PRF"
    }
    Catch {
        $Error[0]
    }
    
    Write-Verbose -Message "Outlook startup is slow. Giving it two seconds before closing"
    Start-Sleep -Seconds 2
    Try {
        Stop-Process -Name "OUTLOOK*"    
    }
    Catch {
        $Error[0]    
    }
}