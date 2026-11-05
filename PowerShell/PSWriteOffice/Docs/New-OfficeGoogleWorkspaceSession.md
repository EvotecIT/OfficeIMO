---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeGoogleWorkspaceSession
## SYNOPSIS
Creates a native Google Workspace session from a service-account key.

## SYNTAX
### ServiceAccount (Default)
```powershell
New-OfficeGoogleWorkspaceSession [-KeyPath] <string> [-Subject <string>] [-Scope <string[]>] [-ApplicationName <string>] [-RequestTimeoutSeconds <int>] [-MaxRetryCount <int>] [<CommonParameters>]
```

## DESCRIPTION
Authenticates the requested scopes through the native JWT credential source. The returned session owns no browser or persistent token store. Dispose its CredentialSource when finished. Service-account PEM import requires PowerShell 7.4 or newer.

## EXAMPLES

### EXAMPLE 1
```powershell
$session = New-OfficeGoogleWorkspaceSession -KeyPath ./Secrets/reporting.json -Subject reports@example.com -Scope Documents,DriveFile
try { Export-OfficeWordGoogleDocument -Path ./Report.docx -Session $session } finally { $session.CredentialSource.Dispose() }
```


## PARAMETERS

### -ApplicationName
Application name sent with requests.

```yaml
Type: String
Parameter Sets: ServiceAccount
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeyPath
Exact service-account JSON key path.

```yaml
Type: String
Parameter Sets: ServiceAccount
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxRetryCount
Maximum native retries.

```yaml
Type: Int32
Parameter Sets: ServiceAccount
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RequestTimeoutSeconds
Per-request timeout.

```yaml
Type: Int32
Parameter Sets: ServiceAccount
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scope
Native scope names or full URLs; default authorizes Docs and Sheets with DriveFile.

```yaml
Type: String[]
Parameter Sets: ServiceAccount
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subject
Delegated user; omission acts as the service account itself.

```yaml
Type: String
Parameter Sets: ServiceAccount
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.GoogleWorkspace.GoogleWorkspaceSession`

## RELATED LINKS

- None
