---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeConfluenceSession
## SYNOPSIS
Creates an in-memory Confluence Cloud session.

## SYNTAX
### Basic (Default)
```powershell
New-OfficeConfluenceSession [-SiteUri] <uri> -Credential <pscredential> [-ApplicationName <string>] [-RequestTimeoutSeconds <int>] [-MaxRetryCount <int>] [<CommonParameters>]
```

### Bearer
```powershell
New-OfficeConfluenceSession [-SiteUri] <uri> -AccessToken <securestring> -CloudId <string> [-ApplicationName <string>] [-RequestTimeoutSeconds <int>] [-MaxRetryCount <int>] [<CommonParameters>]
```

## DESCRIPTION
Creates an in-memory Confluence Cloud session.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $session = New-OfficeConfluenceSession -SiteUri 'https://example.atlassian.net/' -Credential (Get-Credential)
```

Use the Atlassian email as the user name and the API token as the password.

## PARAMETERS

### -AccessToken
OAuth access token stored as a secure string.

```yaml
Type: SecureString
Parameter Sets: Bearer
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ApplicationName
Application name sent with requests.

```yaml
Type: String
Parameter Sets: Basic, Bearer
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CloudId
Atlassian Cloud identifier required for OAuth bearer-token routing.

```yaml
Type: String
Parameter Sets: Bearer
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Credential
Atlassian email and API token credential.

```yaml
Type: PSCredential
Parameter Sets: Basic
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxRetryCount
Maximum retry count for safe read requests. Writes are never retried automatically.

```yaml
Type: Int32
Parameter Sets: Basic, Bearer
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RequestTimeoutSeconds
Per-request timeout in seconds.

```yaml
Type: Int32
Parameter Sets: Basic, Bearer
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SiteUri
HTTPS root URI of the Confluence Cloud site.

```yaml
Type: Uri
Parameter Sets: Basic, Bearer
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Confluence.ConfluenceSession`

## RELATED LINKS

- None
