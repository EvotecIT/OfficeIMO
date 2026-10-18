---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeEmailMailbox
## SYNOPSIS
Reads a native mbox mailbox with bounded per-message diagnostics.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeEmailMailbox [-Path] <string> [-Options <EmailMailboxReaderOptions>] [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Reads a native mbox mailbox with bounded per-message diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeEmailMailboxReaderOptions -MaxMessageCount 5000
Get-OfficeEmailMailbox -Path .\Archive.mbox -Options $options -AsResult
```


## PARAMETERS

### -AsResult
Return the mailbox read result with diagnostics.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional mailbox variant and input limits.

```yaml
Type: EmailMailboxReaderOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to an mbox or mbx file.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Email.EmailMailbox`
- `OfficeIMO.Email.EmailMailboxReadResult`

## RELATED LINKS

- None
