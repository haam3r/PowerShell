#Requires -Version 2.0

<#
.Synopsis
   Convert Unix epoch time
.DESCRIPTION
   Convert Unix epoch time to human readable timestamp.
   Possible inputs are a single Epoch value or CSV file with the fields "Start Time" and "Stop Time"
.EXAMPLE
   Convert-FromUnixDate -Epoch 1508242825
.EXAMPLE
   Convert-FromUnixDate -CSV C:\path\to\file.csv
.EXAMPLE
   Convert-FromUnixDate -CSV .\logs.csv -Verbose -Columns "Start Time","Stop Time"
.EXAMPLE
   Convert-FromUnixDate -CSV C:\path\to\file.csv | Export-Csv .\converted.csv -NoTypeInformation
.EXAMPLE
   $times = 1508242825,1508242825
   foreach ($time in $times) { Convert-FromUnixDate -Epoch $time }
.NOTES
   Author: haam3r
   Version: 1.0
#>
function Convert-FromUnixDate
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0,
                   HelpMessage='Single epoch value',
                   ParameterSetName='SingleValue')]
        [int]$Epoch,

        [Parameter(Mandatory=$false,
                   Position=1,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage='Source CSV file',
                   ParameterSetName='ImportCSV')]
        [ValidateScript({Test-Path $_ -PathType 'Leaf'})]
        [Alias("C")]
        [string]$CSV,

        [Parameter(Mandatory=$false,
                   Position=2,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage='Cloumns to convert',
                   ParameterSetName='ImportCSV')]
        $Columns = @("Start Time","Stop Time")
    )

    Begin {
        # Set Unix beginning of epoch time
        [datetime]$Origin = '1970-01-01 00:00:00'
        
    }

    Process {
        if ($Epoch) {
            Write-Verbose "Got a single epoch time value of $Epoch"
            $Converted = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($Epoch))
            $Report += @(Get-Date -Format "yyyy-MM-dd HH:mm:ss" $Converted)
        }
        elseif ($CSV) {
            Write-Verbose "Importing from $CSV"
            $Logs = Import-Csv -Path $CSV -Delimiter ","

            # Check if specified columns are present in the CSV
            $Headers = $Logs | Get-Member

            foreach($Column in $Columns) {
                if($Headers.Name -contains "$Column") {
                    Write-Verbose "Found Column with Name: $Column"
                }
                else {
                    throw "Could not find column header named: $Column"
                }
            }

            Write-Verbose "Converting timestamps"
            foreach ($Log in $Logs) {
                foreach ($Column in $Columns) {
                    $Log.$Column = Get-Date -Format "yyyy-MM-dd HH:mm:ss" $Origin.AddSeconds($Log.$Column)
                }
                $Report += @($Log)
            }
        }
    }
    End
    {
        Write-Verbose "Processed $($Report.Count) rows"
        $Report
    }
}