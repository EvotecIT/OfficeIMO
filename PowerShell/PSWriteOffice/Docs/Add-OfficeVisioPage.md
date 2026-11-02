---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeVisioPage
## SYNOPSIS
Adds a page to a Visio document and optionally executes nested DSL content.

## SYNTAX
### Document (Default)
```powershell
Add-OfficeVisioPage [-Name] <string> [[-Content] <scriptblock>] [-Document <VisioDocument>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-DisplayUnit <VisioMeasurementUnit>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a page to a Visio document and optionally executes nested DSL content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeVisio -Path .\Workbook.vsdx {
    VisioPage -Name 'Architecture' -Width 11in -Height 8.5in {
        VisioRectangle -Key api -Text 'API' -X 2in -Y 4in
    }
}
```

Adds a named page and executes the nested shape DSL inside that page scope.

## PARAMETERS

### -Content
Nested DSL content executed within this page scope.

```yaml
Type: ScriptBlock
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisplayUnit
Measurement unit displayed by Visio, independent of geometry input units.

```yaml
Type: VisioMeasurementUnit
Parameter Sets: Document
Aliases: None
Possible values: Inches, Centimeters, Millimeters

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Target Visio document. Optional inside New-OfficeVisio.

```yaml
Type: VisioDocument
Parameter Sets: Document
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
Type: OfficeLength
Parameter Sets: Document
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
Parameter Sets: Document
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
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Page width.

```yaml
Type: OfficeLength
Parameter Sets: Document
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
