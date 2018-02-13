# Get-CWCisOnline
## SYNOPSIS
Returns boolean for weather the machine has contacted within the hast $Seconds
## SYNTAX
```powershell
Get-CWCisOnline [-Server] <Object> [-GUID] <Object> [-User] <Object> [-Password] <Object> [[-Seconds] <Object>] [<CommonParameters>]
```
## DESCRIPTION

## PARAMETERS

### -Server &lt;Object&gt;
The address to your Control server example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'
```
Required?                    true
Position?                    1
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -GUID &lt;Object&gt;
The GUID identifier for the machine you wish to connect to.
Cant find documentation on how to find guid but is in url and service
```
Required?                    true
Position?                    2
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -User &lt;Object&gt;
User to authenticate against the control server
```
Required?                    true
Position?                    3
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -Password &lt;Object&gt;
Password to authenticate against the control server
```
Required?                    true
Position?                    4
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -Seconds &lt;Object&gt;
The ammout of time that can have passed for it to be considered online.
Default value is 120 seconds.
```
Required?                    false
Position?                    5
Default value                120
Accept pipeline input?       false
Accept wildcard characters?  false
```
## INPUTS

## OUTPUTS
System.Bool
## NOTES
```
Version:        1.0
Author:         Chris Taylor
Creation Date:  1/20/2016
Purpose/Change: Initial script development
```
## EXAMPLES

### EXAMPLE 1
```powershell
Get-CWCisOnline -Server $Server -GUID $GUID -User $User -Password $Password
```
Will return a true or false value of weather the machine has checked in within the default time (2 minutes)

# Get-CWCLastContact
## SYNOPSIS
Returns the date the machine last connected to the server.
## SYNTAX
```powershell
Get-CWCLastContact [-Server] <Object> [-GUID] <Object> [-User] <Object> [-Password] <Object> [<CommonParameters>]
```
## DESCRIPTION

## PARAMETERS

### -Server &lt;Object&gt;
The address to your Control server example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'
```
Required?                    true
Position?                    1
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -GUID &lt;Object&gt;
The GUID identifier for the machine you wish to connect to.
Cant find documentation on how to find guid but is in url and service
```
Required?                    true
Position?                    2
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -User &lt;Object&gt;
User to authenticate against the control server
```
Required?                    true
Position?                    3
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -Password &lt;Object&gt;
Password to authenticate against the control server
```
Required?                    true
Position?                    4
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
## INPUTS

## OUTPUTS
[datetime]
## NOTES
```
Version:        1.0
Author:         Chris Taylor
Creation Date:  1/20/2016
Purpose/Change: Initial script development
```
## EXAMPLES

### EXAMPLE 1
```powershell
Get-CWCLastContact -Server $Server -GUID $GUID -User $User -Password $Password
```
Will return the last contact of the machine with that GUID

# Invoke-CWCCommand
## SYNOPSIS
Will issue a command against a given machine and return the results.
## SYNTAX
```powershell
Invoke-CWCCommand [-Server] <Object> [-GUID] <Object> [-User] <Object> [-Password] <Object> [[-Command] <Object>] [[-TimeOut] <Object>] [<CommonParameters>]
```
## DESCRIPTION

## PARAMETERS

### -Server &lt;Object&gt;
The address to your Control server example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'
```
Required?                    true
Position?                    1
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -GUID &lt;Object&gt;
The GUID identifier for the machine you wish to connect to.
Cant find documentation on how to find guid but is in url and service
```
Required?                    true
Position?                    2
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -User &lt;Object&gt;
User to authenticate against the control server
```
Required?                    true
Position?                    3
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -Password &lt;Object&gt;
Password to authenticate against the control server
```
Required?                    true
Position?                    4
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -Command &lt;Object&gt;
The command you wish to issue to the machine.
```
Required?                    false
Position?                    5
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -TimeOut &lt;Object&gt;
The amount of time in milliseconds that a command can execute the default is 10 seconds.
```
Required?                    false
Position?                    6
Default value                10000
Accept pipeline input?       false
Accept wildcard characters?  false
```
## INPUTS

## OUTPUTS
The output of the Command provided.
## NOTES
```
Version:        1.0
Author:         Chris Taylor
Creation Date:  1/20/2016
Purpose/Change: Initial script development
```
## EXAMPLES

### EXAMPLE 1
```powershell
Invoke-CWCCommand -Server $Server -GUID $GUID -User $User -Password $Password -Command 'hostname'
```
Will return the hostname of the machine
 
### EXAMPLE 2
```powershell
Invoke-CWCCommand -Server $Server -GUID $GUID -User $User -Password $Password -Command 'powershell "iwr https://bit.ly/ltposh | iex; Restart-LTService"'
```
Will restart the Automate agent on the target machine.

