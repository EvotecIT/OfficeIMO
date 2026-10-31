---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePowerPointInspectionOptions
## SYNOPSIS
Chooses native presentation inspection reports without changing a document.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePowerPointInspectionOptions [-NoPackageValidation] [-NoPreflight] [-NoAccessibility] [-Features] [-ReviewComments] [-Animations] [-Signatures] [-Visuals] [-FileFormatVersion <OfficeOpenXmlFileFormatVersion>] [<CommonParameters>]
```

## DESCRIPTION
Chooses native presentation inspection reports without changing a document.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePowerPointInspectionOptions -Features -ReviewComments
Get-OfficeDocumentInfo -LiteralPath ./Review.pptx -Detailed -Options $options
```


### EXAMPLE 2
```powershell
$options = New-OfficePowerPointInspectionOptions -NoPackageValidation -NoPreflight
(Get-OfficeDocumentInfo -Document $deck -Detailed -Options $options).Details.Accessibility
```


## PARAMETERS

### -Animations
Gets or sets whether to inspect animation metadata.

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

### -Features
Gets or sets whether to inventory supported and preserved features.

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

### -FileFormatVersion
Gets or sets the Open XML version used by package validation; defaults to Microsoft 365.

```yaml
Type: OfficeOpenXmlFileFormatVersion
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Office2007, Office2010, Office2013, Office2016, Office2019, Office2021, Microsoft365

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoAccessibility
Gets or sets whether to omit accessibility inspection.

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

### -NoPackageValidation
Gets or sets whether to omit Open XML package validation.

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

### -NoPreflight
Gets or sets whether to omit layout, text-fit and asset preflight.

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

### -ReviewComments
Gets or sets whether to inspect review comments.

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

### -Signatures
Gets or sets whether to inspect package signature metadata.

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

### -Visuals
Gets or sets whether to create PNG, SVG and snapshot proof for each slide.

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

- `OfficeIMO.PowerPoint.PowerPointInspectionOptions`

## RELATED LINKS

- None
