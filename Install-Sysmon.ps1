#Requires -Version 2.0
function Install-Sysmon {
<#
.Synopsis
   Install Sysmon on multiple machines
.DESCRIPTION
   Install Sysmon, with given config, to any number of machines. Additionaly hide the service. Accepts pipeline input for computer names and has credential support.
   By default installs SwiftOnSecurity's sysmon config.
   Another great option is: "https://raw.githubusercontent.com/ion-storm/sysmon-config/master/sysmonconfig-export.xml"
   Default install path is $env:PROGRAMDATA and installer is downloaded from "https://live.sysinternals.com"
.EXAMPLE
   Install-Sysmon -ComputerName win7x64 -Credential domain\admin
.EXAMPLE
   Get-ADComputer -Filter * | Install-Sysmon -Credential domain\admin
.EXAMPLE
    Install-Sysmon -Credential domain\admin -ComputerName DC,DC2,WIN7X64 -Path "$env:SystemDrive\Programdata\Sysmon" -Config "\\domain.tld\SYSVOL\domain.tld\sysmonconfig-export.xml" -ExeDownload64 "\\domain.tld\SYSVOL\domain.tld\Sysmon64.exe"
.NOTES
   Author: haam3r
   Source: https://github.com/ion-storm/sysmon-config
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage='One or more computer names')]
        [Alias("ComputerName")]
        # Parameter is Name so as to accept pipeline input from the Active Directory PowerShell module
        [string[]]$Name,

        [Parameter(Mandatory=$true,
                   Position=1,
                   #ValueFromPipelineByPropertyName=$true,
                   HelpMessage='Credentials to use')]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory=$false,
                Position=2,
                HelpMessage="Where to put Sysmon. Default is ProgramData\sysmon. Expecting full path.")]
        [Alias("PSPath")]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Path = "$env:SystemDrive\ProgramData\sysmon",

        [Parameter(Mandatory=$false,
                Position=3,
                HelpMessage="Sysmon config file download location. Default is ion-storm's config from github. Expecting URL or UNC path.")]
        [Alias("C")]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Config = "https://raw.githubusercontent.com/SwiftOnSecurity/sysmon-config/master/sysmonconfig-export.xml",

        [Parameter(Mandatory=$false,
                Position=4,
                HelpMessage="Sysmon 64 bit exe download location. Default is live.sysinternals.com. Expecting URL or UNC path.")]
        [Alias("Exe64")]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ExeDownload64 = "https://live.sysinternals.com/Sysmon64.exe",

        [Parameter(Mandatory=$false,
                Position=5,
                HelpMessage="Sysmon 32 bit exe download location. Default is live.sysinternals.com. Expecting URL or UNC path.")]
        [Alias("Exe32")]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ExeDownload32 = "https://live.sysinternals.com/Sysmon.exe"
    )

    BEGIN {
    }

    PROCESS {
        foreach ($Computer in $Name) {
            Write-Output "Installing Sysmon to $Computer"

            Invoke-Command -ComputerName $Computer -Credential $Credential -ArgumentList $Path,$Config,$ExeDownload64,$ExeDownload32 -ScriptBlock {
                
                param($Path,$Config,$ExeDownload64,$ExeDownload32)
                $VerbosePreference=$Using:VerbosePreference
                
                Write-Verbose -Message "Create sysmon directory if needed"

                if (-not (Test-Path $Path)) {
                    New-Item -Path $Path -ItemType Directory
                }

                Set-Location -Path "$Path"

                Write-Verbose -Message "Download the sysmon config file"
                (New-Object System.Net.WebClient).DownloadFile("$Config","$Path\sysmonconfig-export.xml")

                Write-Verbose -Message "Download and install sysmon"
                if ( ((Get-WmiObject Win32_OperatingSystem).OSArchitecture) -eq "64-bit") {
                    (New-Object System.Net.WebClient).DownloadFile("$ExeDownload64","$Path\sysmon64.exe")
                    $SysmonInstall = ".\sysmon64.exe -accepteula -i sysmonconfig-export.xml"
                }
                else {
                    (new-object System.Net.WebClient).DownloadFile("$ExeDownload32","$Path\sysmon.exe")
                    $SysmonInstall = ".\sysmon.exe -accepteula -i sysmonconfig-export.xml"
                }

           }

           Write-Verbose -Message "Set sysmon to restart on service failure"
           Write-Verbose -Message "Hide sysmon from services.msc and Powershell-s Get-Service"
           Hide-Service -ComputerName $Computer -Credential $Credential -ServiceName Sysmon
            
        }
    }
    
    END {
    }

}


function Hide-Service {
<#
.Synopsis
   Hide or reveal service and set failure action to restart 
.DESCRIPTION
   Hide or unhide service by modifying SDDL descriptors on service, additionaly set the service to restart on failure
.EXAMPLE
   Set-LSService -ComputerName test -Credential domain\admin -ServiceName Sysmon
.EXAMPLE
   Set-LSService -ComputerName test -Credential domain\admin -ServiceName Sysmon -Reveal
.NOTES
   Author: haam3r
   Source: https://github.com/ion-storm/sysmon-config
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [Alias("Name")]
        [string]$ComputerName,

        [Parameter(Mandatory=$true,
                   Position=1,
                   HelpMessage='Credentials to use')]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory=$true,
                   Position=2)]
        [Alias("Service")]
        [string]$ServiceName,

        [Parameter(Mandatory=$false,
                   Position=3)]
        [switch]$Reveal
    )

    Begin
    {
        # Define custom output object
        $sddlprops = [ordered]@{'ComputerName' = $ComputerName;
                                'Status' = "";
                                'Comment' = "";}
    }
    Process {
        
        $SetFailure = Invoke-Command -ComputerName $ComputerName -Credential $Credential -ArgumentList $ServiceName -ScriptBlock {
            param($ServiceName)
            $VerbosePreference=$Using:VerbosePreference
            sc.exe failure $ServiceName actions= restart/10000/restart/10000// reset= 120
        }

        if ($Reveal) {
           $sddlset =  Invoke-Command -ComputerName $ComputerName -Credential $Credential -ArgumentList $ServiceName -ScriptBlock {
                param($ServiceName)
                $VerbosePreference=$Using:VerbosePreference
                sc.exe sdset $ServiceName 'D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)'
            }
        }
        else {
            $sddlset = Invoke-Command -ComputerName $ComputerName -Credential $Credential -ArgumentList $ServiceName -ScriptBlock {
                param($ServiceName)
                $VerbosePreference=$Using:VerbosePreference
                sc.exe sdset $ServiceName 'D:(D;;DCLCWPDTSD;;;IU)(D;;DCLCWPDTSD;;;SU)(D;;DCLCWPDTSD;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)'
            }
        }

        if ( $sddlset -like "*FAILED*" ) {
            $sddlprops.Status = $false
            $sddlprops.Comment = $sddlset
            $output = New-Object -TypeName PSObject -Property $sddlprops
            $output
        }
        elseif ( $sddlset -like "*SUCCESS*" ) {
            $sddlprops.Status = $true
            $sddlprops.Comment = $sddlset
            $output = New-Object -TypeName PSObject -Property $sddlprops
            $output
        }
        else {
            $sddlprops.Status = $false
            $sddlprops.Comment = $sddlset
            $output = New-Object -TypeName PSObject -Property $sddlprops
        }
    }
    End {
        return $output
    }
}