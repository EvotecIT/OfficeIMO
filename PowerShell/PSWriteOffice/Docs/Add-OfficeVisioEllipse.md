---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeVisioEllipse
## SYNOPSIS
Adds an ellipse shape to the current Visio page.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeVisioEllipse [[-Text] <string>] [-Page <VisioPage>] [-Key <string>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Name <string>] [-NameU <string>] [-FillColor <OfficeColor>] [-LineColor <OfficeColor>] [-LineWeight <OfficeLength>] [-LinePattern <Int32>] [-FillPattern <Int32>] [-Angle <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an ellipse shape to the current Visio page.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeVisio -Path .\Flow.vsdx {
    VisioEllipse -Key start -Text 'Start' -X 1in -Y 5in -Width 1.4in -Height 0.7in -FillColor '#DCFCE7'
}
```

Adds an ellipse to the active Visio page.

## PARAMETERS

### -Angle
Rotation in degrees.

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

### -FillColor
Fill color or semantic role resolved through the default Office theme.

```yaml
Type: OfficeColor
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FillPattern
Native fill pattern: 0 has no fill, 1 is solid, and 2 through 40 select built-in patterns.

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

### -Height
Physical shape height.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Key
Key used by connector and layout commands within the block.

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

### -LineColor
Line color or semantic role resolved through the default Office theme.

```yaml
Type: OfficeColor
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LinePattern
Native line pattern: 0 hides the line, 1 is solid, and 2 through 23 select built-in patterns.

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

### -LineWeight
Physical outline weight.

```yaml
Type: OfficeLength
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
Optional shape name; otherwise Key is used.

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

### -NameU
Optional universal shape name.

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

### -Page
Target page; the current page is used inside a Visio content block.

```yaml
Type: VisioPage
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
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

### -Text
Text placed inside the shape.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Physical shape width.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Shape pin's X coordinate. A number means points.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Shape pin's Y coordinate. A number means points.

```yaml
Type: OfficeLength
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

- `OfficeIMO.Visio.VisioPage`

## OUTPUTS

- `OfficeIMO.Visio.VisioShape`

## RELATED LINKS

- None
