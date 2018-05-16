<#
.SYNOPSIS
    A PowerShell wrapper for the ConnectWise Control API

.DESCRIPTION
    This module will allow you to interact with the Control API to issue commands and retrieve data.

.NOTES
    Version:        1.0
    Author:         Chris Taylor
    Creation Date:  1/20/2016
    Purpose/Change: Initial script development

.LINK
    labtechconsulting.com
#>

#requires -version 3

#region-[Functions]------------------------------------------------------------

function Get-CWCLastContact {
<#
  .SYNOPSIS
    Returns the date the machine last connected to the server.

  .PARAMETER Server
    The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

  .PARAMETER GUID
    The GUID identifier for the machine you wish to connect to.
    No ddocumentation on how to find the GUID but it is in the URL and service.

  .PARAMETER User
    User to authenticate against the Control server.

  .PARAMETER Password
    Password to authenticate against the Control server.

  .PARAMETER Quiet
    Will output a boolean result, $True for Connected or $False for Offline.

  .PARAMETER Seconds
    Used with the Quiet switch. The number of seconds a machine needs to be offline before returning $False.

  .OUTPUTS
      [datetime]

  .NOTES
      Version:        1.0
      Author:         Chris Taylor
      Creation Date:  1/20/2016
      Purpose/Change: Initial script development

  .EXAMPLE
      Get-CWCLastContact -Server $Server -GUID $GUID -User $User -Password $Password
        Will return the last contact of the machine with that GUID.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        $Server,
        [Parameter(Mandatory=$True)]
        $GUID,
        [Parameter(Mandatory=$True)]
        $User,
        [Parameter(Mandatory=$True)]
        $Password,
        [switch]$Quiet,
        [int]$Seconds
    )

    # Time conversion
    $origin = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0
    $epoch = $((New-TimeSpan -Start $(Get-Date -Date "01/01/1970") -End $(Get-Date)).TotalSeconds)

    $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
    $mycreds = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)

    $Body = @"
    ["All Machines","$GUID"]
"@
    $URl = "$Server/Services/PageService.ashx/GetSessionDetails"
    try {
        $SessionDetails = Invoke-RestMethod -Uri $url -Method Post -Credential $mycreds -ContentType "application/json" -Body $Body
    }
    catch {
        Write-Warning "There was an error connecting to the server."
        Write-Warning "ERROR: $($_.Exception.Message)"
        exit 1
    }

    if ($SessionDetails -eq 'null' -or !$SessionDetails) {
        Write-Warning "Machine not found."
        exit 1
    }

    # Filter to only guest session events
    $GuestSessionEvents = ($SessionDetails.Connections | Where-Object {$_.ProcessType -eq 2}).Events

    if ($GuestSessionEvents) {

        # Get connection events
        $LatestEvent = ($GuestSessionEvents | Where-Object {$_.EventType -in (10,11)} | Sort-Object time)[0]
        if ($LatestEvent.EventType -eq 10) {
            # Currently connected
            if ($Quiet) {
                $True
            } else {
                Get-Date
            }

        }
        else {
            # Time conversion hell :(
            $TimeDiff = $epoch - ($LatestEvent.Time /1000)
            $OfflineTime = $origin.AddSeconds($TimeDiff)
            $Difference = New-TimeSpan –Start $OfflineTime –End $(Get-Date)
            if ($Quiet -and $Difference.Seconds -lt $Seconds) {
                $True
            } elseif ($Quiet -and $Difference.Seconds -gt $Seconds) {
                $False
            } else {
                $OfflineTime
            }
        }
    }
    else {
        Write-Warning "Unable to determine last contact."
        exit 1
    }
}

function Invoke-CWCCommand {
<#
  .SYNOPSIS
    Will issue a command against a given machine and return the results.

  .PARAMETER Server
    The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

  .PARAMETER GUID
    The GUID identifier for the machine you wish to connect to.
    No documentation on how to find the GUID but it is in the URL and service.

  .PARAMETER User
    User to authenticate against the Control server.

  .PARAMETER Password
    Password to authenticate against the Control server.

  .PARAMETER Command
    The command you wish to issue to the machine.

  .PARAMETER TimeOut
    The amount of time in milliseconds that a command can execute. The default is 10000 milliseconds.

  .OUTPUTS
      The output of the Command provided.

  .NOTES
      Version:        1.0
      Author:         Chris Taylor
      Creation Date:  1/20/2016
      Purpose/Change: Initial script development

  .EXAMPLE
      Invoke-CWCCommand -Server $Server -GUID $GUID -User $User -Password $Password -Command 'hostname'
        Will return the hostname of the machine.

  .EXAMPLE
      Invoke-CWCCommand -Server $Server -GUID $GUID -User $User -Password $Password -TimeOut 120000 -Command 'powershell "iwr https://bit.ly/ltposh | iex; Restart-LTService"'
        Will restart the Automate agent on the target machine.
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        $Server,
        [Parameter(Mandatory=$True)]
        $GUID,
        [Parameter(Mandatory=$True)]
        $User,
        [Parameter(Mandatory=$True)]
        $Password,
        $Command,
        $TimeOut = 10000
    )

    $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
    $mycreds = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)

    $origin = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0

    $URI = "$Server/Services/PageService.ashx/AddEventToSessions"
    # Encode the command and create body
    $Command = $Command -replace '(?<!\\)(?:\\)(?!\\)','\\'
    $Command = $Command -replace '"(?<!\\")','\"'
    $Command = @"
#timeout=$TimeOut
$Command
"@
    $Body = @"
["All Machines",["$GUID"],44,"$Command"]
"@

    # Issue command
    try {
        $null = Invoke-RestMethod -Uri $URI -Method Post -Credential $mycreds -ContentType "application/json" -Body $Body
    }
    catch {
        Write-Warning "There was a problem issuing the command."
        Write-Warning "ERROR: $(($_.ErrorDetails | ConvertFrom-Json).message)"
        exit 1
    }

    # Get Session
    $Body = @"
    ["All Machines","$GUID"]
"@
    $URI = "$Server/Services/PageService.ashx/GetSessionDetails"

    try {
        $SessionDetails = Invoke-RestMethod -Uri $URI -Method Post -Credential $mycreds -ContentType "application/json" -Body $Body
    }
    catch {
        Write-Warning "There was a problem validating the command was issued."
        Write-Warning "ERROR: $($_.Exception.Message)"
    }

    #Get time command was executed
    $epoch = $((New-TimeSpan -Start $(Get-Date -Date "01/01/1970") -End $(Get-Date)).TotalSeconds)
    $ExecuteTime = $epoch - ((($SessionDetails.events | Where-Object {$_.EventType -eq 44})[-1]).Time /1000)
    $ExecuteDate = $origin.AddSeconds($ExecuteTime)

    # Look for results of command
    $Looking = $True
    $TimeOut = (Get-Date).AddMilliseconds($TimeOut)
    $Body = @"
["All Machines","$GUID"]
"@
    while ($Looking) {
        try {
            $SessionDetails = Invoke-RestMethod -Uri $URI -Method Post -Credential $mycreds -ContentType "application/json" -Body $Body
        }
        catch {
            Write-Warning "There was a problem validating the command was issued."
            Write-Warning "ERROR: $($_.Exception.Message)"
        }

        $ConnectionsWithData = @()
        Foreach ($Connection in $SessionDetails.connections) {
            $ConnectionsWithData += $Connection | Where-Object {$_.Events.EventType -eq 70}
        }

        $Events = ($ConnectionsWithData.events | Where-Object {$_.EventType -eq 70 -and $_.Time})
        foreach ($Event in $Events) {
            $epoch = $((New-TimeSpan -Start $(Get-Date -Date "01/01/1970") -End $(Get-Date)).TotalSeconds)
            $CheckTime = $epoch - ($Event.Time /1000)
            $CheckDate = $origin.AddSeconds($CheckTime)
            if ($CheckDate -gt $ExecuteDate) {
                $Looking = $False
                $Event.Data -split '[\r\n]' | Where-Object {$_} | Select-Object -skip 1
            }
        }

        Start-Sleep -Seconds 1
        if ($(Get-Date) -gt $TimeOut.AddSeconds(1)) {
            $Looking = $False
        }
    }
}

#endregion Functions
