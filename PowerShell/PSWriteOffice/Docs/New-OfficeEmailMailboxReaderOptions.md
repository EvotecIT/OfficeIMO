---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeEmailMailboxReaderOptions
## SYNOPSIS
Creates bounded mbox reader settings through ordinary PowerShell parameters.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeEmailMailboxReaderOptions [-MessageOptions <EmailReaderOptions>] [-Variant <MboxVariant>] [-MaxMessageCount <int>] [-MaxMailboxBytes <long>] [<CommonParameters>]
```

## DESCRIPTION
Creates bounded mbox reader settings through ordinary PowerShell parameters.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $messageOptions = New-OfficeEmailReaderOptions -ExcludeAttachmentContent
$options = New-OfficeEmailMailboxReaderOptions -MessageOptions $messageOptions -MaxMessageCount 5000
Get-OfficeEmailMailbox -Path .\Archive.mbox -Options $options -AsResult
```


## PARAMETERS

### -MaxMailboxBytes
Maximum aggregate source bytes consumed from one mailbox.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxMessageCount
Maximum messages in one mailbox.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MessageOptions
Bounded policy applied independently to each message.

```yaml
Type: EmailReaderOptions
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
Escaping convention to decode.

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

- `OfficeIMO.Email.EmailReaderOptions`

## OUTPUTS

- `OfficeIMO.Email.EmailMailboxReaderOptions`

## RELATED LINKS

- None
