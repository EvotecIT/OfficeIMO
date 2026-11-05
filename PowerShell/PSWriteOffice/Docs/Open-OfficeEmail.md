---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeEmail
## SYNOPSIS
Reads a native EML, EMLX, MSG, or TNEF artifact with bounded diagnostics.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeEmail [-Path] <string> [-Options <EmailReaderOptions>] [-StoreOptions <EmailStoreReaderOptions>] [-AsResult] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeEmail -LiteralPath <string> [-Options <EmailReaderOptions>] [-StoreOptions <EmailStoreReaderOptions>] [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Reads a native EML, EMLX, MSG, or TNEF artifact with bounded diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeEmailReaderOptions -ExcludeAttachmentContent -MaxAttachmentBytes 25MB
Open-OfficeEmail -Path .\Message.msg -Options $options -AsResult
```


## PARAMETERS

### -AsResult
Return the read result with diagnostics and consumed byte count.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Exact input path without wildcard expansion.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Optional format detection, compound-file, MIME, attachment, and size limits.

```yaml
Type: EmailReaderOptions
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to an EML, EMLX, MSG, TNEF, or winmail.dat file.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -StoreOptions
Optional Apple Mail EMLX envelope, metadata, attachment, and size limits.

```yaml
Type: EmailStoreReaderOptions
Parameter Sets: Path, LiteralPath
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Email.EmailDocument`
- `OfficeIMO.Email.EmailReadResult`
- `OfficeIMO.Email.Store.EmailStoreReadResult`

## RELATED LINKS

- None
