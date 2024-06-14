<#
#>

########
# Global settings
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"
Set-StrictMode -Version 2

<#
#>
Function Select-ForType
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [AllowNull()]
        $Object,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [string]$Type,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [switch]$Derived = $false,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ScriptBlock]$ScriptBlock
    )

    process
    {
        # Make sure we have a type object
        $checkType = [Type]$Type

        # Stop here on null
        if ($null -eq $Object)
        {
            $Object
            return
        }

        # Compare the input object type
        if ($Object.GetType() -eq $checkType -or ($Derived -and $checkType.IsAssignableFrom($Object.GetType())))
        {
            ForEach-Object -InputObject $Object -Process $ScriptBlock
        } else {
            $Object
        }
    }
}

<#
#>
Function Reset-LogFile
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$LogPath,

        [Parameter(Mandatory=$false)]
        [int]$PreserveCount = 5,

        [Parameter(Mandatory=$false)]
        [int]$RotateSizeKB = 0
    )

    process
    {
        # Check if the target is a directory
        if (Test-Path -PathType Container $LogPath)
        {
            Write-Error "Target is a directory"
        }

        # Create the log file, if it doesn't exist
        if (!(Test-Path $LogPath))
        {
            Write-Verbose "Log Path doesn't exist. Attempting to create."
            if ($PSCmdlet.ShouldProcess($LogPath, "Create Log"))
            {
                New-Item -Type File $LogPath -EA SilentlyContinue | Out-Null
            } else {
                return
            }
        }

        # Get the attributes of the target log file
        $logInfo = Get-Item $LogPath
        $logSize = ($logInfo.Length/1024)
        Write-Verbose "Current log file size: $logSize KB"

        # Check if the log file needs rotation
        if ($logSize -lt $RotateSizeKB)
        {
            Write-Verbose "No rotation required"
            return
        }

        # Log file is over the threshold, so need to rotate or truncate
        Write-Verbose "Rotation required due to log size"
        Write-Verbose "PreserveCount: $PreserveCount"

        # Shuffle all of the logs along
        [int]$count = $PreserveCount
        while ($count -gt 0)
        {
            # If count is 1, we're working on the active log
            if ($count -le 1)
            {
                $source = $LogPath
            } else {
                $source = ("{0}.{1}" -f $LogPath, ($count-1))
            }
            $destination = ("{0}.{1}" -f $LogPath, $count)

            # Check if there is an actual log to move and rename
            if (Test-Path -Path $source)
            {
                Write-Verbose "Need to rotate $source"
                if ($PSCmdlet.ShouldProcess($source, "Rotate"))
                {
                    Move-Item -Path $source -Destination $destination -Force
                }
            }

            $count--
        }

        # Create the log path, if it doesn't exist (i.e. was renamed/rotated)
        if (!(Test-Path $LogPath))
        {
            if ($PSCmdlet.ShouldProcess($LogPath, "Create Log"))
            {
                New-Item -Type File $LogPath -EA SilentlyContinue | Out-Null
            } else {
                return
            }
        }

        # Clear the content of the log path (only applies if no rotation was done
        # due to 0 PreserveCount, but the log is over the RotateSizeKB maximum)
        if ($PSCmdlet.ShouldProcess($LogPath, "Truncate"))
        {
            Clear-Content -Path $LogPath -Force
        }
    }
}

<#
#>
Function Format-AsLog
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        $Input
    )

    process
    {
        # Nothing to do for null
        if ($null -eq $Input)
        {
            return
        }

        $timestamp = [DateTime]::Now.ToString("yyyyMMdd HH:mm")

        @($Input) | Select-ForType -Type 'System.String' -Derived -ScriptBlock {
            ("{0} (INFO): {1}" -f $timestamp, $_.ToString())
        } | Select-ForType -Type 'System.Management.Automation.InformationRecord' -Derived -ScriptBlock {
            ("{0} (INFO): {1}" -f $timestamp, $_.ToString())
        } | Select-ForType -Type 'System.Management.Automation.VerboseRecord' -Derived -ScriptBlock {
            ("{0} (VERBOSE): {1}" -f $timestamp, $_.ToString())
        } | Select-ForType -Type 'System.Management.Automation.ErrorRecord' -Derived -ScriptBlock {
            ("{0} (ERROR): {1}" -f $timestamp, $_.ToString())
            $_ | Out-String -Stream | ForEach-Object {
                ("{0} (ERROR): {1}" -f $timestamp, $_.ToString())
            }
        } | Select-ForType -Type 'System.Management.Automation.DebugRecord' -Derived -ScriptBlock {
            ("{0} (DEBUG): {1}" -f $timestamp, $_.ToString())
        } | Select-ForType -Type 'System.Management.Automation.WarningRecord' -Derived -ScriptBlock {
            ("{0} (WARNING): {1}" -f $timestamp, $_.ToString())
        }
    }
}


Function Invoke-ScriptRepeat
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ScriptBlock]$ScriptBlock,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [int]$Iterations = -1,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Start", "Finish")]
        [string]$WaitFrom = "Start",

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [int]$WaitSeconds = 0,

        [Parameter(Mandatory=$false)]
        [switch]$CatchError = $false
    )

    process
    {
        $count = $iterations

        # Iterations:
        # 1+ - Run this many times
        # 0  - Don't run at all
        # -1 - Run indefinitely

        while ($count -gt 0)
        {
            # Only decrement count if it is greater than zero
            # Count could be -1, which implies infinite runs
            # If count is already zero, don't decrement or it
            # will run indefinitely
            if ($count -gt 0)
            {
                $count--
            }

            # Capture start of script run
            $start = [DateTime]::Now
            Write-Verbose ("Start Time: " + $start.ToString("yyyyMMdd HH:mm:ss"))

            # Execute script - Output will be streamed back to the caller
            if ($CatchError)
            {
                try {
                    & $ScriptBlock *>&1
                } catch {
                    # Catch the error and pass it along in the pipeline 
                    $_
                }
            } else {
                & $ScriptBlock *>&1
            }

            # Capture finish time for script run
            $finish = [DateTime]::Now
            Write-Verbose ("Finish Time: " + $finish.ToString("yyyyMMdd HH:mm:ss"))

            # Check if we have runs remaining and determine how
            # to wait
            if ($count -ne 0)
            {
                # Calculate the wait time
                $relative = $finish
                if ($WaitFrom -eq "Start")
                {
                    $relative = $start
                }

                # Determine the wait time in seconds
                $wait = ($relative.AddSeconds($WaitSeconds) - [DateTime]::Now).TotalSeconds
                Write-Verbose "Next iteration in $wait seconds"

                if ($wait -gt 0)
                {
                    # Wait until we should run again
                    Write-Verbose "Starting sleep"
                    Start-Sleep -Seconds $wait
                }
            }
        }
    }
}

