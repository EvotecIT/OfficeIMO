---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeVisioPage
## SYNOPSIS
Adds a page to a Visio document and optionally executes nested DSL content.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeVisioPage [-Name] <string> [[-Content] <scriptblock>] [-Document <VisioDocument>] [-Width <double>] [-Height <double>] [-Unit <VisioMeasurementUnit>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a page to a Visio document and optionally executes nested DSL content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\Workbook.vsdx {
    VisioPage -Name 'Architecture' {
        VisioRectangle -Key api -Text 'API' -X 2 -Y 4
    }
}
```

Adds a named page and executes the nested shape DSL inside that page scope.

## PARAMETERS

### -Content
Nested DSL content executed within this page scope.

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

### -Document
Target Visio document. Optional inside OfficeVisio.

```yaml
Type: VisioDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Height
Page height.

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

### -Name
Page name.

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

### -Unit
Measurement unit for width and height.

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

### -Width
Page width.

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

- `OfficeIMO.Visio.VisioDocument`

## OUTPUTS

- `OfficeIMO.Visio.VisioPage`

## RELATED LINKS

- None
