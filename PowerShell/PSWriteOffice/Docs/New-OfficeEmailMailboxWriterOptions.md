---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeEmailMailboxWriterOptions
## SYNOPSIS
Creates deterministic mbox writer settings through ordinary PowerShell parameters.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeEmailMailboxWriterOptions [-MessageOptions <EmailWriterOptions>] [-Variant <MboxVariant>] [<CommonParameters>]
```

## DESCRIPTION
Creates deterministic mbox writer settings through ordinary PowerShell parameters.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $messageOptions = New-OfficeEmailWriterOptions -IncludeBccHeader
$options = New-OfficeEmailMailboxWriterOptions -MessageOptions $messageOptions -Variant Mboxo
$mailbox | Save-OfficeEmailMailbox -Path .\Archive.mbox -Options $options -PassThru
```


## PARAMETERS

### -MessageOptions
Serialization policy applied independently to each message.

```yaml
Type: EmailWriterOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Variant
Concrete mbox escaping convention to write.

```yaml
Type: MboxVariant
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Auto, Mboxo, Mboxrd

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Email.EmailWriterOptions`

## OUTPUTS

- `OfficeIMO.Email.EmailMailboxWriterOptions`

## RELATED LINKS

- None
