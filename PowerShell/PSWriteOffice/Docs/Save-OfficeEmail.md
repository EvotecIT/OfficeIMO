---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeEmail
## SYNOPSIS
Saves an email document as EML, EMLX, MSG, or TNEF with fidelity diagnostics.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeEmail [-Path] <string> -Document <EmailDocument> [-Format <EmailFileFormat>] [-Options <EmailWriterOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves an email document as EML, EMLX, MSG, or TNEF with fidelity diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeEmailWriterOptions -ConversionLossPolicy Block
$message | Save-OfficeEmail -Path .\Message.eml -Options $options -PassThru
```


## PARAMETERS

### -Document
Email document to save.

```yaml
Type: EmailDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Format
Optional explicit output format. By default it is inferred from the filename.

```yaml
Type: EmailFileFormat
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Unknown, Eml, OutlookMsg, Tnef, Mbox, OutlookTemplate, Emlx

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional preservation, projection, encoding, and output limits.

```yaml
Type: EmailWriterOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

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

### -Path
Destination path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Email.EmailDocument`

## OUTPUTS

- `OfficeIMO.Email.EmailWriteResult`

## RELATED LINKS

- None
