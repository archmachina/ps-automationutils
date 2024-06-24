<#
#>

# Global settings
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"
Set-StrictMode -Version 2

# Global variables
$script:Notifiers = @{}
$script:Automations = @{}

# Classes

Class AutomationUtilsCapture
{
    [System.Collections.Generic.List[System.Object]]$Content

    AutomationUtilsCapture()
    {
        $this.Content = [System.Collections.Generic.List[System.Object]]::New()
    }

    [string]ToString()
    {
        return ($this.Content | Out-String)
    }
}

Class AutomationUtilsNotification
{
    [string]$Title = ""
    [string]$Body = ""
    [string]$Source = ""

    AutomationUtilsNotification()
    {
    }
}

Class AutomationUtilsNotificationBatch
{
    [string]$Name
    [System.Collections.Generic.List[AutomationUtilsNotification]]$Notifications

    AutomationUtilsNotificationBatch([string]$Name)
    {
        $this.Notifications = New-Object 'System.Collections.Generic.List[AutomationUtilsNotification]'
        $this.Name = $Name
    }
}

# Functions

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

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ScriptBlock]$Begin = $null,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ScriptBlock]$Process = $null,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ScriptBlock]$End = $null
    )

    begin
    {
        if ($null -ne $Begin)
        {
            & $Begin
        }

        $objects = New-Object 'System.Collections.Generic.List[System.Object]'
    }

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
            # Only store this object if we have an End script to pass it to
            if ($null -ne $End)
            {
                $objects.Add($Object)
            }

            # Run the process script, if defined
            if ($null -ne $Process)
            {
                ForEach-Object -InputObject $Object -Process $Process
            }
        } else {
            $Object
        }
    }

    end
    {
        if ($null -ne $End)
        {
            ForEach-Object -InputObject $objects -Process $End
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

        @($Input) | Select-ForType -Type 'System.String' -Derived -Process {
            ("{0} (INFO): {1}" -f $timestamp, $_.ToString())
        } | Select-ForType -Type 'System.Management.Automation.InformationRecord' -Derived -Process {
            ("{0} (INFO): {1}" -f $timestamp, $_.ToString())
        } | Select-ForType -Type 'System.Management.Automation.VerboseRecord' -Derived -Process {
            ("{0} (VERBOSE): {1}" -f $timestamp, $_.ToString())
        } | Select-ForType -Type 'System.Management.Automation.ErrorRecord' -Derived -Process {
            ("{0} (ERROR): {1}" -f $timestamp, $_.ToString())
            $_ | Out-String -Stream | ForEach-Object {
                ("{0} (ERROR): {1}" -f $timestamp, $_.ToString())
            }
        } | Select-ForType -Type 'System.Management.Automation.DebugRecord' -Derived -Process {
            ("{0} (DEBUG): {1}" -f $timestamp, $_.ToString())
        } | Select-ForType -Type 'System.Management.Automation.WarningRecord' -Derived -Process {
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
                    Write-Verbose "Error received from script block"
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

<#
#>
Function New-Capture
{
    [CmdletBinding()]
    param()

    process
    {
        [AutomationUtilsCapture]::New()
    }
}

<#
#>
Function Copy-ToCapture
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [AutomationUtilsCapture]$Capture,

        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [AllowNull()]
        $Object
    )

    process
    {
        $Capture.Content.Add($Object)
        $Object
    }
}

<#
#>
Function Invoke-CaptureScript
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [AutomationUtilsCapture]$Capture,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ScriptBlock]$ScriptBlock
    )

    process
    {
        & $ScriptBlock *>&1 | Copy-ToCapture -Capture $capture
    }
}

Function New-Notification
{
    [CmdletBinding(DefaultParameterSetName="Body")]
    param(
        [Parameter(Mandatory=$true, ParameterSetName="Body")]
        [Parameter(Mandatory=$true, ParameterSetName="Script")]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Mandatory=$true, ParameterSetName="Body")]
        [ValidateNotNullOrEmpty()]
        [string]$Body = $null,

        [Parameter(Mandatory=$true, ParameterSetName="Script")]
        [ValidateNotNull()]
        [ScriptBlock]$ScriptBlock = $null,

        [Parameter(Mandatory=$false, ParameterSetName="Body")]
        [Parameter(Mandatory=$false, ParameterSetName="Script")]
        [ValidateNotNullOrEmpty()]
        [string]$Source = ""
    )

    process
    {
        $notification = [AutomationUtilsNotification]::New()
        $notification.Title = $Title
        $notification.Source = $Source

        switch ($PSCmdlet.ParameterSetName) {
            "Body" {
                $notification.Body = $Body
            }
            "Script" {
                $capture = New-Capture
                Invoke-CaptureScript -Capture $capture -ScriptBlock $ScriptBlock
                $notification.Body = $capture.ToString()
            }
            default {
                Write-Error "Unknown parameter set name"
            }
        }

        $notification
    }
}

Function Send-Notifications
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [AllowNull()]
        $Notification,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [switch]$Pass,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    begin
    {
        $batch = [AutomationUtilsNotificationBatch]::New($Name)
    }

    process
    {
        # Check if we have a notification object to record
        if ($null -ne $Notification -and ([AutomationUtilsNotification].IsAssignableFrom($Notification.GetType())))
        {
            $batch.Notifications.Add($Notification)
            return
        }

        # If requested, pass the object on in the pipeline, rather than error
        if ($Pass)
        {
            $Notification
            return
        }

        # Not a notification object and we're not requested to pass it on, so error
        Write-Error "Non-Notification passed to Send-Notifications and 'Pass' not enabled"
    }

    end
    {
        # Check if we have some notifications to send
        if (($batch.Notifications | Measure-Object).Count -lt 1)
        {
            # Nothing to do
            return
        }

        # Iterate by Notifier as each notifier should receive a batch of the notifications
        # to allow it to choose whether to batch send or send individually
        $script:Notifiers.Keys | ForEach-Object {
            $notifierName = $_
            $notifier = $script:Notifiers[$notifierName]

            Write-Verbose ("Sending notification via " + $notifierName)
            ForEach-Object -InputObject $batch -Process $notifier
        }
    }
}

Function Register-Notifier
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Name="default",

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ScriptBlock]$ScriptBlock
    )

    process
    {
        $script:Notifiers[$Name] = $ScriptBlock
    }
}

Function Register-Automation
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ScriptBlock]$ScriptBlock
    )

    process
    {
        $script:Automations[$Name] = $ScriptBlock
    }
}

Function Invoke-Automation
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [HashTable]$Config = @{}
    )

    process
    {
        # Make sure this automation exists
        if (!$script:Automations.ContainsKey($Name))
        {
            Write-Error "Automation `"$Name`" does not exist"
        }

        $automation = $script:Automations[$Name]

        # Run the automation
        & $automation @Config *>&1 | Select-ForType -Type AutomationUtilsNotification -Derived -Process {
            $_.Source = $Name
            $_
        }
    }
}

<#
#>
Function Limit-StringLength
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [ValidateNotNull()]
        [string]$Str,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [int]$Length
    )

    process
    {
        $working = $Str

        # Going to include ..., so must be 4 or higher
        if ($Length -lt 4)
        {
            Write-Error "Length must be 4 or greater"
        }

        # Truncate and add '...' if the length is too high
        if ($working.Length -gt $Length)
        {
            $working = $working.Substring(0, $Length-3) + "..."
        }

        $working
    }
}

<#
#>
Function Split-StringLength
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [AllowEmptyString()]
        [ValidateNotNull()]
        [string]$Str,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [int]$WrapLength
    )

    process
    {
        if ($WrapLength -lt 1)
        {
            Write-Error "Invalid WrapLength - Must be greater than 0."
        }

        $working = $Str

        while ($working.Length -gt $WrapLength)
        {
            $newStr = $working.Substring(0, $WrapLength)
            $working = $working.Substring($WrapLength, $working.Length-$WrapLength)

            $newStr
        }

        $working
    }
}

