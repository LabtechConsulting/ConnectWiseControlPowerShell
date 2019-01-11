This is a PowerShell wrapper for the ConnectWise Control API.
https://docs.connectwise.com/ConnectWise_Control_Documentation/Developers/Session_Manager_API_Reference


```
irm 'https://bit.ly/controlposh' | iex
```

example:
```
# Your Control server URL
$Server = 'https://control.domain.com'

# Get Control credentials
$Credentials = Get-Credential

# Load module into memory
irm 'https://bit.ly/controlposh' | iex

# Find this machine in Control
$Computer = Get-CWCSessions -Server $Server -Credentials $Credentials -Type Access -Search $env:COMPUTERNAME -Limit 1

# Get the machines last contact
Get-CWCLastContact -Server $Server -Credentials $Credentials -GUID $Computer.SessionID
```
         
         
# Functions

[Get-CWCLastContact](CWCPoSh/Get-CWCLastContact.md)

[Get-CWCSessions](CWCPoSh/Get-CWCSessions.md)

[Invoke-CWCCommand](CWCPoSh/Invoke-CWCCommand.md)

[Invoke-CWCWake](CWCPoSh/Invoke-CWCWake.md)

[Remove-CWCSession](CWCPoSh/Remove-CWCSession.md)

[Update-CWCSessionName](CWCPoSh/Update-CWCSessionName.md)


