---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeVisioContainer
## SYNOPSIS
Creates an OfficeIMO-authored Visio-native container around existing shapes.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeVisioContainer [[-InputObject] <Object>] -Id <string> [-Page <VisioPage>] [-ShapeId <string[]>] [-Text <string>] [-Margin <Double>] [-HeadingHeight <Double>] [-FillColor <string>] [-LineColor <string>] [-LineWeight <Double>] [-ContainerStyle <Int32>] [-HeadingStyle <Int32>] [-Locked] [-NoAutoResize] [-NoHighlight] [-NoRibbon] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates an OfficeIMO-authored Visio-native container around existing shapes.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\Architecture.vsdx {
    VisioRectangle -Key api -Text 'API' -X 2 -Y 4
    VisioRectangle -Key worker -Text 'Worker' -X 4 -Y 4
    VisioContainer -Id app -Text 'Application tier' -ShapeId api,worker -FillColor '#F8FAFC'
}
```

Creates a native Visio container around previously keyed shapes.

## PARAMETERS

### -ContainerStyle
Native Visio container gallery identifier. Zero selects the default/no explicit gallery style.

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

### -FillColor
Container fill color.

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

### -HeadingHeight
Additional heading height in page units.

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

### -HeadingStyle
Native Visio heading gallery identifier. Zero disables the heading style.

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

### -Id
Container shape identifier.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Shapes, shape selections, or shape keys/ids to include in the container.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LineColor
Container line color.

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

### -LineWeight
Container line weight in inches.

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

### -Locked
Lock the generated container.

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

### -Margin
Outer margin around member shapes in page units.

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

### -NoAutoResize
Disable Visio container auto resize metadata.

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

### -NoHighlight
Suppress Visio selection highlighting metadata.

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

### -NoRibbon
Suppress Visio container ribbon metadata.

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

### -Page
Page that owns the member shapes. Optional inside OfficeVisio/VisioPage DSL scopes.

```yaml
Type: VisioPage
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

### -ShapeId
Shape keys or ids to include in the container.

```yaml
Type: String[]
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
Container heading text.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `OfficeIMO.Visio.VisioShape`

## RELATED LINKS

- None
