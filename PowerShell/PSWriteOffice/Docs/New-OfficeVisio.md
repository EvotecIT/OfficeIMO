---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeVisio
## SYNOPSIS
Creates a new OfficeIMO.Visio document with an initial page and optional DSL content.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeVisio [-Path] <string> [[-Content] <scriptblock>] [-PageName <string>] [-Width <double>] [-Height <double>] [-Unit <VisioMeasurementUnit>] [-Title <string>] [-Author <string>] [-RequestRecalcOnOpen] [-UseMastersByDefault] [-NoSave] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Creates a new OfficeIMO.Visio document with an initial page and optional DSL content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\ServiceMap.vsdx -Title 'Service map' -RequestRecalcOnOpen {
    VisioRectangle -Key web -Text 'Web' -X 1 -Y 4 -FillColor LightBlue
    VisioRectangle -Key api -Text 'API' -X 4 -Y 4 -FillColor LightGreen
    VisioConnector -From web -To api -EndArrow Triangle -Label 'calls'
}
```

Creates an editable .vsdx diagram with two shapes and a connector.

## PARAMETERS

### -Author
Optional document author.

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

### -Content
DSL script block describing Visio pages, shapes, and connectors.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Initial page height.

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

### -NoSave
Skip saving and emit the document object.

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
Open the document after saving.

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
Initial page name.

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
Emit the document object instead of the saved file.

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
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RequestRecalcOnOpen
Ask Visio to recalculate layout and connector routing when the document opens.

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

### -Title
Optional document title.

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

### -Unit
Measurement unit for page width and height.

```yaml
Type: VisioMeasurementUnit
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Inches, Centimeters, Millimeters

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseMastersByDefault
Use Visio masters for supported built-in stencil shapes when saving.

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

### -Width
Initial page width.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Visio.VisioDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
