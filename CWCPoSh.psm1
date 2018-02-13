<#
.SYNOPSIS
    A powershell wrapper for the ConnectWise Control API

.DESCRIPTION
    This module will allow you to interact with the Control API allowing you to retreive data and issue commands.

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
    The address to your Control server example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

  .PARAMETER GUID
    The GUID identifier for the machine you wish to connect to.
    Cant find documentation on how to find guid but is in url and service

  .PARAMETER User
    User to authenticate against the control server

  .PARAMETER Password
    Password to authenticate against the control server

  .PARAMETER Quite
    Will output a boolean result $True for Connected $False for Offline

  .PARAMETER Seconds
    Used with Quite switch. The number of secconds a machine needs to be offline before returning $False

  .OUTPUTS
      [datetime]

  .NOTES
      Version:        1.0
      Author:         Chris Taylor
      Creation Date:  1/20/2016
      Purpose/Change: Initial script development

  .EXAMPLE
      Get-CWCLastContact -Server $Server -GUID $GUID -User $User -Password $Password
        Will return the last contact of the machine with that GUID
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Server,
        [Parameter(Mandatory=$true)]
        $GUID,
        [Parameter(Mandatory=$true)]
        $User,
        [Parameter(Mandatory=$true)]
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
        Write-Warning "There was was an error connecting to the server."
        Write-Warning "ERROR: $($_.Exception.Message)"
        exit 1
    }

    if ($SessionDetails -eq 'null' -or !$SessionDetails) {
        Write-Warning "Machine not found."
        exit 1
    }

    # Filter to guest session events
    $GuestSessionEvents = ($SessionDetails.Connections | Where-Object{$_.ProcessType -eq 2}).Events

    if ($GuestSessionEvents) {

        # Get connection events
        $LatestEvent = ($GuestSessionEvents |  Where-Object{$_.EventType -in (10,11)} | Sort-Object time)[0]
        if($LatestEvent.EventType -eq 10){
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
                $Difference.Seconds
                
            } elseif ($Quiet -and $Difference.Seconds -gt $Seconds) {
                $False
                $Difference.Seconds
            } else {
                $OfflineTime
            }
        }
    }
    else {
        Write-Warning "Unable to determin last contact."
        exit 1
    }
}

function Invoke-CWCCommand {
<#
  .SYNOPSIS
    Will issue a command against a given machine and return the results.

  .PARAMETER Server
    The address to your Control server example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

  .PARAMETER GUID
    The GUID identifier for the machine you wish to connect to.
    Cant find documentation on how to find guid but is in url and service

  .PARAMETER User
    User to authenticate against the control server

  .PARAMETER Password
    Password to authenticate against the control server

  .PARAMETER Command
    The command you wish to issue to the machine.

  .PARAMETER TimeOut
    The amount of time in milliseconds that a command can execute the default is 10 seconds.

  .OUTPUTS
      The output of the Command provided.

  .NOTES
      Version:        1.0
      Author:         Chris Taylor
      Creation Date:  1/20/2016
      Purpose/Change: Initial script development

  .EXAMPLE
      Invoke-CWCCommand -Server $Server -GUID $GUID -User $User -Password $Password -Command 'hostname'
        Will return the hostname of the machine

  .EXAMPLE
      Invoke-CWCCommand -Server $Server -GUID $GUID -User $User -Password $Password -Command 'powershell "iwr https://bit.ly/ltposh | iex; Restart-LTService"'
        Will restart the Automate agent on the target machine.
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Server,
        [Parameter(Mandatory=$true)]
        $GUID,
        [Parameter(Mandatory=$true)]
        $User,
        [Parameter(Mandatory=$true)]
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
        Write-Warning "There was a problem validating command was issued."
        Write-Warning "ERROR: $($_.Exception.Message)"
    }

    #Get time command was executed
    $epoch = $((New-TimeSpan -Start $(Get-Date -Date "01/01/1970") -End $(Get-Date)).TotalSeconds)
    $ExecuteTime = $epoch - ((($SessionDetails.events | Where-Object {$_.EventType -eq 44})[-1]).Time /1000)
    $ExecuteDate = $origin.AddSeconds($ExecuteTime)

    # Look for results of command
    $Looking = $true
    $TimeOut = (Get-Date).AddMilliseconds($TimeOut)
    $Body = @"
["All Machines","$GUID"]
"@
    while($Looking){
        try {
            $SessionDetails = Invoke-RestMethod -Uri $URI -Method Post -Credential $mycreds -ContentType "application/json" -Body $Body
        }
        catch {
            Write-Warning "There was a problem validating command was issued."
            Write-Warning "ERROR: $($_.Exception.Message)"
        }

        $ConnectionsWithData = @()
        Foreach($Connection in $SessionDetails.connections){
            $ConnectionsWithData += $Connection | Where-Object {$_.Events.EventType -eq 70}
        }

        $Events = ($ConnectionsWithData.events | Where-Object {$_.EventType -eq 70 -and $_.Time})
        foreach($Event in $Events) {
            $epoch = $((New-TimeSpan -Start $(Get-Date -Date "01/01/1970") -End $(Get-Date)).TotalSeconds)
            $CheckTime = $epoch - ($Event.Time /1000)
            $CheckDate = $origin.AddSeconds($CheckTime)
            if($CheckDate -gt $ExecuteDate){
                $Looking = $false
                $Event.Data -split '[\r\n]' | Where-Object {$_} | Select-Object -skip 1
            }
        }

        Start-Sleep -Seconds 1
        if($(Get-Date) -gt $TimeOut.AddSeconds(1)){
            $Looking = $false
        }
    }
}

#endregion Functions
