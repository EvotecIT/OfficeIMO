---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeEmailWriterOptions
## SYNOPSIS
Creates deterministic email writer settings through ordinary PowerShell parameters.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeEmailWriterOptions [-ConversionLossPolicy <EmailConversionLossPolicy>] [-UsePreservedRawSource] [-IncludeBccHeader] [-Base64LineLength <int>] [-MaxNestedMessageDepth <int>] [-MaxOutputBytes <long>] [<CommonParameters>]
```

## DESCRIPTION
Creates deterministic email writer settings through ordinary PowerShell parameters.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeEmailWriterOptions -UsePreservedRawSource -ConversionLossPolicy Block
$message | Save-OfficeEmail -Path .\Message.eml -Options $options -PassThru
```


## PARAMETERS

### -Base64LineLength
Maximum encoded characters on one Base64 body line.

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

### -ConversionLossPolicy
Policy applied when the requested format cannot preserve known message semantics.

```yaml
Type: EmailConversionLossPolicy
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Block, Warn, Allow

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeBccHeader
Write Bcc recipients into the message header.

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

### -MaxNestedMessageDepth
Maximum embedded-message write depth.

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

### -MaxOutputBytes
Maximum serialized artifact size.

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

### -UsePreservedRawSource
Emit an unchanged preserved source instead of regenerating the artifact when possible.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Email.EmailWriterOptions`

## RELATED LINKS

- None
