# Invoke-CWCWake
## SYNOPSIS
Will issue a wake command to a given session.
## SYNTAX
```powershell
Invoke-CWCWake -Server <String> -GUID <Guid[]> -User <String> -Password <String> -Type <String> [<CommonParameters>]



Invoke-CWCWake -Server <String> -GUID <Guid[]> -Type <String> -Credentials <PSCredential> [<CommonParameters>]
```
## DESCRIPTION
Will issue a wake command to a given access or support session.
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
### -GUID &lt;Guid[]&gt;
The GUID identifier for the session you wish to end. Accepts an array of GUIDs
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
### -Type &lt;String&gt;
The type of session Support/Access
```
Required                    true
Position                    named
Default value
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
## EXAMPLES
### EXAMPLE 1
```powershell
PS C:\>End-CWWake -Server $Server -GUID $GUID -User $User -Password $Password

Will issue a wake command to a given session.
```

## NOTES
Version:        1.0

Author:         Chris Taylor

Creation Date:  12/7/2018

Purpose/Change: Initial script development 
