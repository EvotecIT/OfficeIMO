---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeVisioGallery
## SYNOPSIS
Generates the OfficeIMO Visio reference gallery as editable .vsdx diagrams.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeVisioGallery [-OutputDirectory] <string> [-NoPackageValidation] [-NoVisualQualityAnalysis] [<CommonParameters>]
```

## DESCRIPTION
Generates the OfficeIMO Visio reference gallery as editable .vsdx diagrams.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeVisioGallery -OutputDirectory .\VisioGallery |
                Select-Object Name, FilePath, IsClean
```

Creates polished, editable Visio samples for flowcharts, architecture, network, timeline, swimlane, org chart, and graph diagrams.

## PARAMETERS

### -NoPackageValidation
Skip structural package validation after gallery documents are generated.

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

### -NoVisualQualityAnalysis
Skip visual quality analysis after gallery documents are generated.

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

### -OutputDirectory
Directory that receives generated .vsdx gallery documents.

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

- `None`

## OUTPUTS

- `OfficeIMO.Visio.VisioGalleryResult`

## RELATED LINKS

- None
