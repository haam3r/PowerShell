#Requires -Version 2.0
function Install-Winlogbeat {
<#
.SYNOPSIS
    Install winlogbeat
.DESCRIPTION
    Deploy the winlogbeat log forwarding solution to multiple machines. Install as a service, with config and hide the service.
.EXAMPLE
    Install-Winlogbeat -ComputerName win7x64 -Credential domain\admin
.EXAMPLE
    Get-ADComputer -Filter * | Install-Winlogbeat -Credential domain\admin
.NOTES
   Author: haam3r
#>
    [CmdletBinding()]
    [Alias()]
    Param (
        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage='One or more computer names')]
        [Alias("ComputerName")]
        # Parameter naming is Name so as to accept pipeline input from the Active Directory PowerShell module
        [string[]]$Name,

        [Parameter(Mandatory=$false,
                   Position=1,
                   HelpMessage='Credentials to use')]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory=$false,
                   Position=2,
                   HelpMessage="Where to place winlogbeat. Default is ProgramData\winlogbeat. Expecting full path.")]
        [Alias("PSPath")]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Path = "$env:ProgramData\winlogbeat",

        [Parameter(Mandatory=$false,
                   Position=3,
                   HelpMessage="Winlogbeat config file download location. Default is C:\Files\winlogbeat.yml. Expecting local path.")]
        [Alias("C")]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Config = "C:\Files\winlogbeat.yml",

        [Parameter(Mandatory=$false,
                   Position=4,
                   HelpMessage="Winlogbeat 64-bit exe download location. Default is C:\Files\winlogbeat64.exe. Expecting local path.")]
        [Alias("Exe64")]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ExeDownload64 = "C:\Files\winlogbeat64.exe",

        [Parameter(Mandatory=$false,
                Position=5,
                HelpMessage="Winlogbeat 32-bit exe download location. Default is C:\Files\winlogbeat32.exe. Expecting local path.")]
        [Alias("Exe32")]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ExeDownload32 = "C:\Files\winlogbeat32.exe"
    )

    Begin {
    }

    Process {
         foreach ($Computer in $Name) {
            Write-Output "Installing Winlogbeat to $Computer"

            
            $OSInfo = Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock {
                $Arch = Get-WmiObject Win32_OperatingSystem
                $Version = [System.Environment]::OSVersion.Version
                $Properties = @{Arch = $Arch.OSArchitecture;
                                MajorVersion = $Version.Major;
                                MinorVersion = $Version.Minor;}
                $Output = New-Object -TypeName PSObject -Property $Properties
                $Output
            }
            New-Item -Path "\\$Computer\C$\ProgramData\winlogbeat" -ItemType Directory -ErrorAction SilentlyContinue
            Copy-Item -Path $Config -Destination "\\$Computer\C$\ProgramData\winlogbeat\winlogbeat.yml" -Force

            if ( $OSInfo.Arch -eq "64-bit") {
                Write-verbose "Copying $ExeDownload64 to $Computer at $Path"
                Copy-Item -Path "$ExeDownload64" -Destination "\\$Computer\C$\ProgramData\winlogbeat\winlogbeat.exe" -Force
            }
            else {
                Write-Verbose "Copying $ExeDownload32 to $Computer at $Path"
                Copy-Item -Path "$ExeDownload32" -Destination "\\$Computer\C$\ProgramData\winlogbeat\winlogbeat.exe" -Force
            }

            Invoke-Command -ComputerName $Computer -Credential $Credential -ArgumentList $Path,$Config,$OSInfo -ScriptBlock {
                param($Path,$Config,$OSInfo)
                $VerbosePreference=$Using:VerbosePreference
                
                Set-Location -Path "$Path"
                Write-Verbose -Message "Checking if service exists and deleting if it does"
                if (Get-Service winlogbeat -ErrorAction SilentlyContinue) {
                    $service = Get-WmiObject -Class Win32_Service -Filter "name='winlogbeat'"
                    $service.StopService()
                    Start-Sleep -Seconds 1
                    $service.delete()
                }
                
                Write-Verbose -Message "Creating winlogbeat service"
                New-Service -Name winlogbeat -DisplayName winlogbeat -BinaryPathName "`"$Path\\winlogbeat.exe`" -c `"$Path\\winlogbeat.yml`" -path.home `"$Path`" -path.data `"C:\\ProgramData\\winlogbeat`""
                Get-Service -Name winlogbeat | Start-Service
            }
        }
    }

    End {
    }
}