---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeOpenDocumentTextBox
## SYNOPSIS
Adds a positioned text box to an OpenDocument presentation slide.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeOpenDocumentTextBox [-Text] <string> [-Slide <OdpSlide>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Name <string>] [-PassThru] [-Bold] [-Italic] [-Strikethrough] [-Underline <OdfTextDecorationStyle>] [-Font <string>] [-Size <OfficeLength>] [-Color <OfficeColor>] [-Alignment <OfficeAlignment>] [<CommonParameters>]
```

## DESCRIPTION
Adds a positioned text box to an OpenDocument presentation slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeOpenDocumentTextBox -Text 'Approved' -X 18cm -Y 12cm -Width 6cm -Height 2cm
```


## PARAMETERS

### -Alignment
Paragraph horizontal alignment.

```yaml
Type: OfficeAlignment
Parameter Sets: Default
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Apply bold text; explicit false clears it.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Opaque RGB text color, including CSS names and theme roles.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Font family.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Height as a physical length; a number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Apply italic text; explicit false clears it.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Optional shape name.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the created text box.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Font size; a number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Slide target. Omit inside Add-OfficeOpenDocumentSlide -Content.

```yaml
Type: OdpSlide
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Strikethrough
Apply strikethrough; explicit false clears it.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text box content.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Native ODF underline line style.

```yaml
Type: OdfTextDecorationStyle
Parameter Sets: Default
Aliases: None
Possible values: None, Solid, Dotted, Dash, LongDash, DotDash, DotDotDash, Wave

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width as a physical length; a number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Horizontal position as a physical length; a number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Vertical position as a physical length; a number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
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

- `OfficeIMO.OpenDocument.OdpSlide`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdpTextBox`

## RELATED LINKS

- None
