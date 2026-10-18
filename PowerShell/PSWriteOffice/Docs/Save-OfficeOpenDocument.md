---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeOpenDocument
## SYNOPSIS
Saves a native OpenDocument model with entry-level preservation diagnostics.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeOpenDocument [-Path] <string> -Document <OdfDocument> [-Options <OdfSaveOptions>] [-FailOnLoss] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves a native OpenDocument model with entry-level preservation diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
Save-OfficeOpenDocument -Path 'C:\Path' -Document 'Value'
```


## PARAMETERS

### -Document
OpenDocument model to save.

```yaml
Type: OdfDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FailOnLoss
Throw when source entries cannot be preserved losslessly.

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
Optional package save and preservation settings.

```yaml
Type: OdfSaveOptions
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
Emit the save result, including preservation diagnostics.

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

- `OfficeIMO.OpenDocument.OdfDocument`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdfSaveResult`

## RELATED LINKS

- None
