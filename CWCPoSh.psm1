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

    Update Date:  1/11/2019
    Purpose/Change: Move to [PSCredential] authentication

.LINK
    labtechconsulting.com
#>

#requires -version 3

if([Net.SecurityProtocolType]::Tls) {
    [Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls
}

#region-[Functions]------------------------------------------------------------

function Get-CWCLastContact {
    <#
      .SYNOPSIS
        Returns the date the machine last connected to the control server.

      .DESCRIPTION
        Returns the date the machine last connected to the control server.

      .PARAMETER Server
        The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

      .PARAMETER GUID
        The GUID/SessionID for the machine you wish to connect to.
        You can retrieve session info with the 'Get-CWCSessions' commandlet

        On Windows clients, the launch parameters are located in the registry at:
          HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\ScreenConnect Client (xxxxxxxxxxxxxxxx)\ImagePath
        On Linux and Mac clients, it's found in the ClientLaunchParameters.txt file in the client installation folder:
          /opt/screenconnect-xxxxxxxxxxxxxxxx/ClientLaunchParameters.txt

      .PARAMETER Credentials
        [PSCredential] object used to authenticate against Control.

      .PARAMETER User
        User to authenticate against the Control server.

      .PARAMETER Password
        Password to authenticate against the Control server.

      .PARAMETER Quiet
        Will output a boolean result, $True for Connected or $False for Offline.

      .PARAMETER Seconds
        Used with the Quiet switch. The number of seconds a machine needs to be offline before returning $False.

      .PARAMETER Group
        Name of session group to use.

      .OUTPUTS
          [datetime] -or [boolean]

      .NOTES
          Version:        1.1
          Author:         Chris Taylor
          Creation Date:  1/20/2016
          Purpose/Change: Initial script development

          Update Date:  8/24/2018
          Purpose/Change: Fix Timespan Seconds duration

      .EXAMPLE
          Get-CWCLastContact -Server $Server -GUID $GUID -User $User -Password $Password
            Will return the last contact of the machine with that GUID.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True)]
        [guid]$GUID,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$User,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$Password,
        [switch]$Quiet,
        [int]$Seconds,
        [string]$Group = "All Machines",
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials
    )

    # Time conversion
    $origin = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0
    $epoch = $((New-TimeSpan -Start $(Get-Date -Date "01/01/1970") -End $(Get-Date)).TotalSeconds)

    if($Password) {
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)
        Write-Warning "Switch to -Credentials [PSCredential] authentication method."
    }

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    $Body = ConvertTo-Json @(@($Group),$GUID) -Compress
    Write-Verbose $Body
    
    $URl = "$Server/Services/PageService.ashx/GetSessionDetails"
    try {
        $SessionDetails = Invoke-RestMethod -Uri $url -Method Post -ContentType "application/json" -Body $Body -Headers $Headers 
    }
    catch {
        Write-Error "$($_.Exception.Message)"
        return
    }

    if ($SessionDetails -eq 'null' -or !$SessionDetails) {
        Write-Error "Machine not found."
        return $null
    }

    # Filter to only guest session events
    $GuestSessionEvents = ($SessionDetails.Connections | Where-Object {$_.ProcessType -eq 2}).Events

    if ($GuestSessionEvents) {

        # Get connection events
        $LatestEvent = ($GuestSessionEvents | Where-Object {$_.EventType -in (10,11)} | Sort-Object time)[0]
        if ($LatestEvent.EventType -eq 10) {
            # Currently connected
            if ($Quiet) {
                return $True
            } else {
                return Get-Date
            }

        }
        else {
            # Time conversion hell :(
            $TimeDiff = $epoch - ($LatestEvent.Time /1000)
            $OfflineTime = $origin.AddSeconds($TimeDiff)
            $Difference = New-TimeSpan -Start $OfflineTime -End $(Get-Date)
            if ($Quiet -and $Difference.TotalSeconds -lt $Seconds) {
                return $True
            } elseif ($Quiet) {
                return $False
            } else {
                return $OfflineTime
            }
        }
    }
    else {
        Write-Error "Unable to determine last contact."
        return
    }
}

function Invoke-CWCCommand {
    <#
    .SYNOPSIS
        Will issue a command against a given machine and return the results.

    .DESCRIPTION
        Will issue a command against a given machine and return the results.

    .PARAMETER Server
        The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

    .PARAMETER GUID
        The GUID identifier for the machine you wish to connect to.
        You can retrieve session info with the 'Get-CWCSessions' commandlet

    .PARAMETER Credentials
        [PSCredential] object used to authenticate against Control.

    .PARAMETER User
        User to authenticate against the Control server.

    .PARAMETER Password
        Password to authenticate against the Control server.

    .PARAMETER Command
        The command you wish to issue to the machine.

    .PARAMETER TimeOut
        The amount of time in milliseconds that a command can execute. The default is 10000 milliseconds.

    .PARAMETER PowerShell
        Issues the command in a powershell session.

    .PARAMETER Group
        Name of session group to use.

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
        Invoke-CWCCommand -Server $Server -GUID $GUID -User $User -Password $Password -TimeOut 120000 -Command 'iwr -UseBasicParsing "https://bit.ly/ltposh" | iex; Restart-LTService' -PowerShell
            Will restart the Automate agent on the target machine.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True)]
        [guid]$GUID,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$User,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$Password,
        [string]$Command,
        [int]$TimeOut = 10000,
        [switch]$PowerShell,
        [string]$Group = "All Machines",
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials
    )

    if($Password) {
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)
        Write-Warning "Switch to -Credentials [PSCredential] authentication method."
    }

    $origin = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0

    $URI = "$Server/Services/PageService.ashx/AddEventToSessions"

    # Format command
    $FormattedCommand = @()
    if ($Powershell) {
        $FormattedCommand += '#!ps'
    }
    $FormattedCommand += "#timeout=$TimeOut"
    $FormattedCommand += $Command
    $FormattedCommand = $FormattedCommand | Out-String

    $SessionEventType = 44
    $Body = (ConvertTo-Json @($Group,@($GUID),$SessionEventType,$FormattedCommand)).Replace('\r\n','\n')
    Write-Verbose $Body
    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    # Issue command
    try {
        $null = Invoke-RestMethod -Uri $URI -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
    }
    catch {
        Write-Error $_
        return
    }

    # Get Session
    $URI = "$Server/Services/PageService.ashx/GetSessionDetails"
    $Body = ConvertTo-Json @($Group,$GUID)
    Write-Verbose $Body
    try {
        $SessionDetails = Invoke-RestMethod -Uri $URI -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
    }
    catch {
        Write-Error $_
        return
    }

    #Get time command was executed
    $epoch = $((New-TimeSpan -Start $(Get-Date -Date "01/01/1970") -End $(Get-Date)).TotalSeconds)
    $ExecuteTime = $epoch - ((($SessionDetails.events | Where-Object {$_.EventType -eq 44})[-1]).Time /1000)
    $ExecuteDate = $origin.AddSeconds($ExecuteTime)

    # Look for results of command
    $Looking = $True
    $TimeOutDateTime = (Get-Date).AddMilliseconds($TimeOut)
    $Body = ConvertTo-Json @($Group,$GUID)
    while ($Looking) {
        try {
            $SessionDetails = Invoke-RestMethod -Uri $URI -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
        }
        catch {
            Write-Error $_
            return
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
                $Output = $Event.Data -split '[\r\n]' | Where-Object {$_ -and $_ -ne "C:\WINDOWS\system32>$Command"}
                Write-Verbose $Event.Data
                return $Output
            }
        }

        Start-Sleep -Seconds 1
        if ($(Get-Date) -gt $TimeOutDateTime.AddSeconds(1)) {
            $Looking = $False
            Write-Warning "Command timed out."
        }
    }
}

function Get-CWCSessions {
    <#
    .SYNOPSIS
        Will return a list of sessions.

    .DESCRIPTION
        Allows you to search for access or service sessions.

    .PARAMETER Server
        The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

    .PARAMETER Credentials
        [PSCredential] object used to authenticate against Control.

    .PARAMETER User
        User to authenticate against the Control server.

    .PARAMETER Password
        Password to authenticate against the Control server.

    .PARAMETER Type
        The type of session Support/Access

    .PARAMETER Group
        Name of session group to use.

    .PARAMETER Search
        Limit results with search pattern.

    .PARAMETER Limit
        Limit the number of results returned.

    .OUTPUTS
        ConnectWise Control session objects

    .NOTES
        Version:        1.0
        Author:         Chris Taylor
        Creation Date:  10/10/2018
        Purpose/Change: Initial script development

    .EXAMPLE
        Get-CWCAccessSessions -Server $Server -User $User -Password $Password -Search "server1" -Limit 10
        Will return the first 10 access sessions that match 'server1'.

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$User,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$Password,
        [Parameter(Mandatory=$True)]
        [ValidateSet('Support','Access')]
        $Type,
        [string]$Group = "All Machines",
        [string]$Search,
        [int]$Limit,
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials
    )

    if($Password) {
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)
        Write-Warning "Switch to -Credentials [PSCredential] authentication method."
    }

    $URI = "$Server/Services/PageService.ashx/GetHostSessionInfo"

    switch($Type){
        'Support'   {$Number = 0}
        'Meeting'   {$Number = 1}
        'Access'    {$Number = 2}
        default     {Write-Error "Unknown Type, $Type";return}
    }

    $Body = ConvertTo-Json @($Number,@($Group),$Search,$null,$Limit)
    Write-Verbose $Body

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    try {
        $Data = Invoke-RestMethod -Uri $URI -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
        return $Data.sessions
    }
    catch {
        Write-Error $_
        return
    }
}

function Remove-CWCSession {
    <#
    .SYNOPSIS
        Will end a given session.

    .DESCRIPTION
        Will end a given access or support session.

    .PARAMETER Server
        The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

    .PARAMETER Credentials
        [PSCredential] object used to authenticate against Control.

    .PARAMETER User
        User to authenticate against the Control server.

    .PARAMETER Password
        Password to authenticate against the Control server.

    .PARAMETER Type
        The type of session Support/Meeting/Access

    .PARAMETER GUID
        The GUID identifier for the session you wish to end. Accepts an array of GUIDs.

    .NOTES
        Version:        1.0
        Author:         Chris Taylor
        Creation Date:  10/10/2018
        Purpose/Change: Initial script development

    .EXAMPLE
        Remove-CWCAccessSession -Server $Server -GUID $GUID -User $User -Password $Password
            Will remove the given access session
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True)]
        [guid[]]$GUID,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$User,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$Password,
        [Parameter(Mandatory=$True)]
        [ValidateSet('Access','Meeting','Support')]
        $Type,
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials
    )

    if($Password) {
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)
        Write-Warning "Switch to -Credentials [PSCredential] authentication method."
    }

    $URI = "$Server/Services/PageService.ashx/AddEventToSessions"

    switch($Type){
        'Support'   {$Group = 'All Sessions'}
        'Access'    {$Group = 'All Machines'}
        default     {Write-Error "Unknown Type, $Type";return}
    }

    $SessionEventType = 21
    if($GUID.count -eq 1){
        $Body = ConvertTo-Json @(@($Group),@($GUID),$SessionEventType,'')
    }
    else {
        $Body = ConvertTo-Json @(@($Group),$GUID,$SessionEventType,'')
    }
    Write-Verbose $Body

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    # Issue command
    try {
        $null = Invoke-RestMethod -Uri $URI -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
    }
    catch {
        Write-Error $(($_.ErrorDetails | ConvertFrom-Json).message)
        return
    }
}

function Update-CWCSessionName {
    <#
    .SYNOPSIS
    Updates the name of a session.

    .DESCRIPTION
        Updates the name of a session on the control server.

    .PARAMETER Server
    The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

    .PARAMETER GUID
    The GUID/SessionID for the machine you wish to connect to.
    You can retrieve session info with the 'Get-CWCSessions' commandlet

    On Windows clients, the launch parameters are located in the registry at:
        HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\ScreenConnect Client (xxxxxxxxxxxxxxxx)\ImagePath
    On Linux and Mac clients, it's found in the ClientLaunchParameters.txt file in the client installation folder:
        /opt/screenconnect-xxxxxxxxxxxxxxxx/ClientLaunchParameters.txt

    .PARAMETER Credentials
    [PSCredential] object used to authenticate against Control.

    .PARAMETER User
    User to authenticate against the Control server.

    .PARAMETER Password
    Password to authenticate against the Control server.

    .PARAMETER NewName
    The new name for the session.

    .NOTES
        Version:        1.1
        Author:         Chris Taylor
        Creation Date:  10/25/2018
        Purpose/Change: Initial script development

    .EXAMPLE
        Update-CWCSessionName -Server $Server -GUID $GUID -User $User -Password $Password -NewName 'Session1'
        Will rename the session to Session1
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True)]
        [guid]$GUID,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$User,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$Password,
        [Parameter(Mandatory=$True)]
        [string]$NewName,
        [string]$Group = "All Machines",
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials
    )

    if($Password) {
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)
        Write-Warning "Switch to -Credentials [PSCredential] authentication method."
    }

    $Body = ConvertTo-Json @($Group,$GUID,$NewName)
    Write-Verbose $Body

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    $URl = "$Server/Services/PageService.ashx/UpdateSessionName"
    try {
        $null = Invoke-RestMethod -Uri $url -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
    }
    catch {
        Write-Error $_.Exception.Message
        return
    }
}

function Invoke-CWCWake {
    <#
      .SYNOPSIS
          Will issue a wake command to a given session.

      .DESCRIPTION
          Will issue a wake command to a given access or support session.

      .PARAMETER Server
          The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

      .PARAMETER Credentials
          [PSCredential] object used to authenticate against Control.

      .PARAMETER User
          User to authenticate against the Control server.

      .PARAMETER Password
          Password to authenticate against the Control server.

      .PARAMETER Type
          The type of session Support/Access

      .PARAMETER GUID
          The GUID identifier for the session you wish to end. Accepts an array of GUIDs

      .NOTES
          Version:        1.0
          Author:         Chris Taylor
          Creation Date:  12/7/2018
          Purpose/Change: Initial script development

      .EXAMPLE
          End-CWWake -Server $Server -GUID $GUID -User $User -Password $Password
            Will issue a wake command to a given session.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True)]
        [guid[]]$GUID,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$User,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$Password,
        [Parameter(Mandatory=$True)]
        [ValidateSet('Support','Access')]
        [string]$Type,
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials
    )

    if($Password) {
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)
        Write-Warning "Switch to -Credentials [PSCredential] authentication method."
    }

    $URI = "$Server/Services/PageService.ashx/AddEventToSessions"

    switch($Type){
        'Support'   {$Group = 'All Sessions'}
        'Access'    {$Group = 'All Machines'}
        default     {Write-Error "Unknown Type, $Type";return}
    }

    $SessionEventType = 43
    if($GUID.count -eq 1){
        $Body = ConvertTo-Json @($Group,@($GUID),$SessionEventType,'')
    }
    else {
        $Body = ConvertTo-Json @($Group,$GUID,$SessionEventType,'')
    }

    Write-Verbose $Body

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    # Issue command
    try {
        $null = Invoke-RestMethod -Uri $URI -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
    }
    catch {
        Write-Error $_
        return
    }
}

function Get-CWCSessionDetail {
    <#
    .SYNOPSIS
        Will return information about a session.

    .DESCRIPTION
        Dispays more information about a session.

    .PARAMETER Server
        The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

    .PARAMETER Credentials
        [PSCredential] object used to authenticate against Control.

    .PARAMETER User
        User to authenticate against the Control server.

    .PARAMETER Password
        Password to authenticate against the Control server.

    .PARAMETER Group
        Name of session group to use.

    .PARAMETER GUID
        GUID of the machine to retreive session details.

    .OUTPUTS
        ConnectWise Control session objects

    .NOTES
        Version:        1.0
        Author:         Chris Taylor
        Creation Date:  1/15/2019
        Purpose/Change: Initial script development

    .EXAMPLE
        Get-CWCAccessSessions -Server $Server -User $User -Password $Password -Search "server1" -Limit 10
        Will return the first 10 access sessions that match 'server1'.

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$User,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$Password,
        [string]$Group = "All Machines",
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials,
        [Parameter(Mandatory=$True)]
        [guid]$GUID
    )

    if($Password) {
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)
        Write-Warning "Switch to -Credentials [PSCredential] authentication method."
    }

    $URI = "$Server/Services/PageService.ashx/GetSessionDetails"

    $Body = ConvertTo-Json @($Group,$GUID)
    Write-Verbose $Body

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    try {
        $Data = Invoke-RestMethod -Uri $URI -Method Post -Headers $Headers -ContentType "application/json; charset=utf-8" -Body $Body -Verbose
        return $Data
    }
    catch {
        Write-Error $_
        return
    }
}

function Update-CWCCustomProperty {
    <#
    .SYNOPSIS
    Updated the custom .

    .DESCRIPTION
        Updates the name of a session on the control server.

    .PARAMETER Server
    The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'

    .PARAMETER GUID
    The GUID/SessionID for the machine you wish to connect to.
    You can retrieve session info with the 'Get-CWCSessions' commandlet

    On Windows clients, the launch parameters are located in the registry at:
        HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\ScreenConnect Client (xxxxxxxxxxxxxxxx)\ImagePath
    On Linux and Mac clients, it's found in the ClientLaunchParameters.txt file in the client installation folder:
        /opt/screenconnect-xxxxxxxxxxxxxxxx/ClientLaunchParameters.txt

    .PARAMETER Credentials
    [PSCredential] object used to authenticate against Control.

    .PARAMETER User
    User to authenticate against the Control server.

    .PARAMETER Password
    Password to authenticate against the Control server.

    .PARAMETER NewName
    The new name for the session.

    .NOTES
        Version:        1.1
        Author:         Chris Taylor
        Creation Date:  10/25/2018
        Purpose/Change: Initial script development

    .EXAMPLE
        Update-CWCSessionName -Server $Server -GUID $GUID -User $User -Password $Password -NewName 'Session1'
        Will rename the session to Session1
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True)]
        [guid]$GUID,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$User,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$Password,
        [Parameter(Mandatory=$True)]
        [string]$NewName,
        [string]$Group = "All Machines",
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials
    )

    if($Password) {
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)
        Write-Warning "Switch to -Credentials [PSCredential] authentication method."
    }

    $Body = ConvertTo-Json @($Group,$GUID,$NewName)
    Write-Verbose $Body

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    $URl = "$Server/Services/PageService.ashx/UpdateSessionName"
    try {
        $null = Invoke-RestMethod -Uri $url -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
    }
    catch {
        Write-Error $_
        return
    }
}

function New-CWCUser {
    <#
    .SYNOPSIS
    Creates a user.

    .DESCRIPTION
    Creates a new local user.

    .PARAMETER Server
    The ConnectWise Control server you care connecting to

    .PARAMETER Credentials
    Credentials of the non MFA user used to perform the actions.

    .PARAMETER UserName
    Username of the new user

    .PARAMETER Password
    Password of the new user

    .PARAMETER OPT
    The MFA token

    .PARAMETER DisplayName
    The display name of the new user

    .PARAMETER SecurityGroups
    An array of security groups the user is a part of

    .PARAMETER ForcePassChange
    Force the user to change their password at next login

    .NOTES
        Version:        1.1
        Author:         Chris Taylor
        Creation Date:  4/2/2020
        Purpose/Change: Initial script development

    .EXAMPLE

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials,
        [Parameter(Mandatory=$True)]
        [string]$UserName,
        [Parameter(Mandatory=$True)]
        [string]$Password,
        [string]$OTP,
        [string]$DisplayName,
        [string]$Email,
        [string[]]$SecurityGroups,
        [boolean]$ForcePassChange = $true
    )

    $Body = ConvertTo-Json -Depth 10 @(
        "XmlMembershipProvider",
        $null,
        $UserName,
        $Password,
        $Password,
        $OTP,
        $DisplayName,
        "",
        $Email,
        $SecurityGroups,
        $ForcePassChange
    )
    Write-Verbose $Body

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    $URl = "$Server/Services/SecurityService.ashx/SaveUser"
    $null = Invoke-RestMethod -Uri $url -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
}

function New-CWCAssignment {
    <#
    .SYNOPSIS
    Assign a user to machines.

    .DESCRIPTION
    Uses the remote workforce extension to assign a user to machines

    .PARAMETER Server

    .PARAMETER GUID

    .PARAMETER Credentials

    .PARAMETER User

    .PARAMETER Password

    .PARAMETER NewName

    .NOTES
        Version:        1.1
        Author:         Chris Taylor
        Creation Date:  10/25/2018
        Purpose/Change: Initial script development

    .EXAMPLE

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials,
        [Parameter(Mandatory=$True)]
        [guid[]]$GUID,
        [Parameter(Mandatory=$True)]
        [string]$Username,
        [string]$DisplayName
    )

    $Body = ConvertTo-Json @(
        $GUID,
        "UserName:$($Username),UserDisplayName:$($DisplayName)"
    )
    Write-Verbose $Body

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    $URl = "$Server/App_Extensions/2c4f522f-b39a-413a-8807-dc52a2fce13e/Service.ashx/AddAssignmentNoteToSession"
    $null = Invoke-RestMethod -Uri $url -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
}

function createRequiredRemoteWorkforceRole {
    <#
    .SYNOPSIS
    Create a new security role

    .DESCRIPTION
    Create a new security role

    .PARAMETER Server

    .PARAMETER GUID

    .PARAMETER Credentials

    .PARAMETER User

    .PARAMETER Password

    .PARAMETER NewName

    .NOTES
        Version:        1.1
        Author:         Chris Taylor
        Creation Date:  10/25/2018
        Purpose/Change: Initial script development

    .EXAMPLE

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials,
        $SessionGroups = @('My Assigned Machines'),
        $RoleName = 'Remote Workforce'
    )

    $Body = ConvertTo-Json -Depth 10 @(
        "",
        $RoleName,
        @(),
        @(
            @{
                "AccessControlType" = 0
                "Name" = "ViewSessionGroup"
                "SessionGroupFilter" = 7
                "SessionGroupPath" = $SessionGroups
                "OwnershipFilter" = 0
            },
            @{
                "AccessControlType" = 0
                "Name" = "JoinSession"
                "SessionGroupFilter" = 7
                "SessionGroupPath" = $SessionGroups
                "OwnershipFilter" = 0
            },
            @{
                "AccessControlType" = 0
                "Name" = "HostSessionWithoutConsent"
                "SessionGroupFilter" = 7
                "SessionGroupPath" = $SessionGroups
                "OwnershipFilter" = 0
            }
        )
    )
    Write-Verbose $Body

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    $URl = "$Server/Services/SecurityService.ashx/SaveRole"
    $null = Invoke-RestMethod -Uri $url -Method Post -Headers $Headers -ContentType "application/json" -Body $Body
}

function New-CWCMFA {
    $Possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"
    $Key = ""
    while ($Key.Length -lt 16) {
        $Key += $Possible.ToCharArray() | Get-Random
    }

    $googleUrlLabel = "otpauth://totp/screenconnect?secret=$Key"
    $qrUrl = "https://chart.googleapis.com/chart?cht=qr&chs=300x300&chl=$($googleUrlLabel)&chld=H|0"
    [pscustomobject]@{
        'QR' = $qrUrl
        'OTP' = "ms:$Key"
    }
}

function Get-CWCSecurityConfigurationInfo {
    <#
    .SYNOPSIS
        Will return security configuration information.

    .DESCRIPTION
        Will return security configuration information.

    .PARAMETER Server
        The address to your Control server. Example 'https://control.christaylor.codes' or 'http://control.secure.me:8040'

    .PARAMETER Credentials
        PSCredential object used to authenticate against Control.

    .PARAMETER User
        User to authenticate against the Control server.

    .PARAMETER Password
        Password to authenticate against the Control server.

    .PARAMETER Group
        Name of session group to use.

    .PARAMETER GUID
        GUID of the machine to retrieve session details.

    .OUTPUTS
        ConnectWise Control session objects

    .NOTES
        Version:        1.0
        Author:         Chris Taylor
        Creation Date:  1/15/2019
        Purpose/Change: Initial script development

    .EXAMPLE
        Get-CWCAccessSessions -Server $Server -User $User -Password $Password -Search "server1" -Limit 10
        Will return the first 10 access sessions that match 'server1'.

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [string]$Server,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$User,
        [Parameter(Mandatory=$True, ParameterSetName='password')]
        [string]$Password,
        [string]$Group = "All Machines",
        [Parameter(Mandatory=$True, ParameterSetName='cred')]
        [PSCredential]$Credentials
    )

    if($Password) {
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($User, $secpasswd)
        Write-Warning "Switch to -Credentials [PSCredential] authentication method."
    }

    $URI = "$Server/Services/SecurityService.ashx/GetSecurityConfigurationInfo"

    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Credentials.UserName):$($Credentials.GetNetworkCredential().Password)")) 
    $Headers = @{ Authorization = "Basic $encodedCredentials" }

    $Data = Invoke-RestMethod -Uri $URI -Method Post -Headers $Headers -ContentType "application/json; charset=utf-8"
    return $Data
}

#endregion Functions
