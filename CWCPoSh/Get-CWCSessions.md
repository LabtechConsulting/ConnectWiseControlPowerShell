# Get-CWCSessions
## SYNOPSIS
Will return a list of sessions.
## SYNTAX
```powershell
Get-CWCSessions -Server <String> -User <String> -Password <String> -Type <Object> [-Group <String>] [-Search <String>] [-Limit <Int32>] [<CommonParameters>]



Get-CWCSessions -Server <String> -Type <Object> [-Group <String>] [-Search <String>] [-Limit <Int32>] -Credentials <PSCredential> [<CommonParameters>]
```
## DESCRIPTION
Allows you to search for access or service sessions.
## PARAMETERS
### -Server &lt;String&gt;
The address to your Control server. Example 'https://control.labtechconsulting.com' or 'http://control.secure.me:8040'
```
Required                    true
Position                    named
Default value
Accept pipeline input       false
Accept wildcard characters  false
```
### -User &lt;String&gt;
User to authenticate against the Control server.
```
Required                    true
Position                    named
Default value
Accept pipeline input       false
Accept wildcard characters  false
```
### -Password &lt;String&gt;
Password to authenticate against the Control server.
```
Required                    true
Position                    named
Default value
Accept pipeline input       false
Accept wildcard characters  false
```
### -Type &lt;Object&gt;
The type of session Support/Access
```
Required                    true
Position                    named
Default value
Accept pipeline input       false
Accept wildcard characters  false
```
### -Group &lt;String&gt;
Name of session group to use.
```
Required                    false
Position                    named
Default value                All Machines
Accept pipeline input       false
Accept wildcard characters  false
```
### -Search &lt;String&gt;
Limit results with search patern.
```
Required                    false
Position                    named
Default value
Accept pipeline input       false
Accept wildcard characters  false
```
### -Limit &lt;Int32&gt;
Limit the number of results returned.
```
Required                    false
Position                    named
Default value                0
Accept pipeline input       false
Accept wildcard characters  false
```
### -Credentials &lt;PSCredential&gt;
[PSCredential] object used to authenticate against Control.
```
Required                    true
Position                    named
Default value
Accept pipeline input       false
Accept wildcard characters  false
```
## OUTPUTS
ConnectWise Control session objects

## EXAMPLES
### EXAMPLE 1
```powershell
PS C:\>Get-CWCAccessSessions -Server $Server -User $User -Password $Password -Search "server1" -Limit 10

Will return the first 10 access sessions that match 'server1'.
```

## NOTES
Version:        1.0

Author:         Chris Taylor

Creation Date:  10/10/2018

Purpose/Change: Initial script development 
