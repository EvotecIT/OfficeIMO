---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeVisioVisual
## SYNOPSIS
Exports CFX semantic visual-artifact input as a native editable VSDX diagram.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficeVisioVisual [-InputObject] <Object> [-Path] <string> [-Open] [-PassThru] [-PageName <string>] [-UseNaturalPageSize] [-PixelsPerInch <double>] [-NoTitle] [-NoGroups] [-NoShapeData] [-NoHyperlinks] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports CFX semantic visual-artifact input as a native editable VSDX diagram.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $topology | ConvertTo-ImageVisualArtifact | Export-OfficeVisioVisual -Path .\Topology.vsdx
```

Creates native Visio shapes, containers, connectors, Shape Data, and hyperlinks from the portable CFX semantics.

## PARAMETERS

### -InputObject
Typed CFX artifact, semantic envelope/JSON, ImagePlayground portable artifact, or prior conversion result.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoGroups
Do not create native Visio containers for CFX groups or lanes.

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

### -NoHyperlinks
Do not copy safe CFX links onto native Visio shapes and connectors.

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

### -NoShapeData
Do not copy CFX metadata, ports, and details into Visio Shape Data.

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

### -NoTitle
Do not add the artifact title as an editable Visio title.

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

### -Open
Open the generated VSDX after saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageName
Name of the generated Visio page.

```yaml
Type: String
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
Emit the conversion result instead of the saved file.

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
Destination .vsdx path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PixelsPerInch
Pixel density used with -UseNaturalPageSize.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseNaturalPageSize
Use the CFX natural pixel size as the minimum Visio page size.

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

- `System.Object`

## OUTPUTS

- `OfficeIMO.ChartForgeX.OfficeVisioVisualConversionResult`
- `System.IO.FileInfo`

## RELATED LINKS

- None
