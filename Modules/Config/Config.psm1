<#
    .SYNOPSIS
        Reads a Config file and return an object
#>

Function Get-Config {
    <#
        .SYNOPSIS
            Reads a configuration file and returns a configuration object to be used in methods
            The base class loads the default configuration. If you need a different configuration in a script
            the simply call this method again or replace a member of the $config object
        .EXAMPLE
            $config = Config -Configfile 'config_localhost.cfg'
    #>
    [CmdletBinding()]
    param(
        # The filename of the configuration file, you can keep multiple config file e.g. local / dev / prod
        # and easily switch by calling a different configuration
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelinebyPropertyName = $True)]
        [string]
        $path
    )

    Begin {
        $config = New-Object -TypeName PSObject
    }

    Process {
        Get-Content $path | convertfrom-stringdata | ForEach-Object {
            foreach ($key in $_.keys) {
                $config | Add-Member -Type NoteProperty -Name $key -Value $_[$key] -Force
            }
            $config | Add-Member -Type NoteProperty -Name 'user' -Value $env:username -Force
            $config | Add-Member -Type NoteProperty -Name 'userprofile' -Value $env:userprofile -Force
        }
    }

    End {
        return $config
    }
}