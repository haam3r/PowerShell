#Requires -Version 2.0

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
function Set-OutlookContact {
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Source csv file from wich to extract user list.
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   HelpMessage="Source csv file from wich to extract user list")]
        [ValidateScript({Test-Path $_ -PathType 'Leaf'})]
        [Alias("C")]
        [string]$CSV
    )

    Begin
    {
        # Import CSV file and setup the com object for Outlook communication
        Write-Verbose "Importing from $CSV"
        $contacts = Import-Csv $CSV -Delimiter ";"
        Write-Verbose "Creating com object to interact with Outlook"
        $outlook = new-object -com Outlook.Application -ea 1
        Write-Verbose "Setting up a contacts session"
        $DefaultFolder = $outlook.session.GetDefaultFolder(10)
    }

    Process
    {
        #Parse the list and import contacts
        Write-Verbose "Adding mail contacts to Outlook"
        $contacts | ForEach-Object {
            $newcontact = $DefaultFolder.Items.Add()
            foreach ($property in $_.PSObject.Properties) {
                $newcontact.$($property.Name) = $property.Value
            }
            $newcontact.Save()    
            }
    }

    End
    {
        # Notify that processing has ended
        $count = ($contacts | Measure-Object).Count
        Write-Verbose "Added $count contacts"
        # Release com objects and close Outlook
        Write-Verbose "Close Outlook connection and remove com object"
        $outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
    }

}