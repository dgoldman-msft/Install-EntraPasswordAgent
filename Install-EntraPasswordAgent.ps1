function Install-EntraPasswordAgent {
    <#
    .SYNOPSIS
        Install the Entra Password Protection Agent

    .DESCRIPTION
        Installs and configures the Microsoft Entra Password Protection Agent to work with your on-prem DC(s)

    .PARAMETER LoggingPath
        Default log file path

    .PARAMETER LoggingFile
        Default log file name

    .EXAMPLE
        PS C:\> Install-EntraPasswordAgent

        Installs the Microsoft Entra Password Protection Agent

    .NOTES
       https://aka.ms/aadpasswordprotectiondocs
       https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-password-ban-bad-on-premises-deploy

    #>

    [CmdletBinding()]
    [Alias([InstallEPP])]
    [OutputType([System.String])]

    param(
        [string]
        $LoggingPath = "C:\EntraPWLogging\",

        [string]
        $LoggingFile = "EntraPWAgentLog.txt"
    )

    begin {
        Write-Output 'Starting Entra Password Protection Agent configuration process'
        Write-Verbose 'Saving current execution policy and changing new policy setting to: Bypass'
        $oldExecutionPolicy = Get-ExecutionPolicy
        Set-ExecutionPolicy -ExecutionPolicy Bypass -ErrorAction Stop
    }

    process {

        try {
            Start-Transcript -Path (Join-Path -Path $LoggingPath -ChildPath $LoggingFile) -Append
            Write-Output "Output  and agent extraction path: $($env:TEMP)"
            $outpath = "$env:TEMP\AzureADPasswordProtectionDCAgentSetup.msi"
            $url = 'https://download.microsoft.com/download/9/7/0/970F8006-9599-4D5A-AA4A-0C66A3A78FE8/AzureADPasswordProtectionDCAgentSetup.msi'
            Write-Output "Checking to see if Entra Password Protection Agent has already been downloaded to $($env:ComputerName)"

            if (-NOT (Test-Path -Path "$env:TEMP\AzureADPasswordProtectionDCAgentSetup.msi")) {
                Write-Output "Not found! Downloading Entra Password Protection Agent to $($env:ComputerName)"
                Invoke-WebRequest -Uri $url -OutFile $outpath
            }
            else {
                Write-Output 'Entra Password Protection Agent has been previously downloaded!'
            }
        }
        catch {
            "Error: $_"
        }

        try {
            Write-Output "Extracting Entra Password Protection Agent package contents to $($env:TEMP). Installing Entra Password Protection Agent on local machine."
            Start-Process -Filepath "$env:TEMP\AzureADPasswordProtectionDCAgentSetup.msi" -ArgumentList "/quiet /norestart /t:$($env:TEMP)" -Wait -ErrorAction Stop
            Write-Output 'Installation Complete!'
        }
        catch {
            "Error: $_"
        }
    }

    end {
        Write-Verbose "Reversing execution policy settings"
        Set-ExecutionPolicy -ExecutionPolicy $oldExecutionPolicy -ErrorAction Stop
        Write-Output "Completed Entra Password Protection Agent agent configuration process. Logging saved to $(Join-Path -Path $LoggingPath -ChildPath $LoggingFile)"
    }
}