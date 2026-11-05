---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfCanvasText
## SYNOPSIS
Adds plain or rich native text to the active PDF canvas.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePdfCanvasText [-Text] <string> -X <OfficeLength> -Y <OfficeLength> [-Width <OfficeLength>] [-Height <OfficeLength>] [-Alignment <OfficeAlignment>] [-LineHeight <OfficeLength>] [-Bold] [-Italic] [-Strikethrough] [-Underline <OfficeTextDecorationStyle>] [-Color <OfficeColor>] [-BackgroundColor <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-Baseline <PdfTextBaseline>] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePdfCanvasText [-Content] <scriptblock> -X <OfficeLength> -Y <OfficeLength> [-Width <OfficeLength>] [-Height <OfficeLength>] [-Alignment <OfficeAlignment>] [-LineHeight <OfficeLength>] [-Color <OfficeColor>] [-Size <OfficeLength>] [<CommonParameters>]
```

## DESCRIPTION
Adds plain or rich native text to the active PDF canvas.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePdfCanvasText -Text 'Value' -X 'Value' -Y 'Value'
```


### EXAMPLE 2
```powershell
Add-OfficePdfCanvasText -Content { } -X 'Value' -Y 'Value'
```


## PARAMETERS

### -Alignment
Alignment for this content.

```yaml
Type: OfficeAlignment
Parameter Sets: Text, Content
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor
BackgroundColor for this content.

```yaml
Type: OfficeColor
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Baseline
Baseline for this content.

```yaml
Type: PdfTextBaseline
Parameter Sets: Text
Aliases: None
Possible values: Normal, Superscript, Subscript

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Bold for this content.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Color for this content.

```yaml
Type: OfficeColor
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Content for this content.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Font for this content.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Height for this content.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Italic for this content.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineHeight
LineHeight for this content.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Size for this content.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Strikethrough for this content.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text for this content.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Underline for this content.

```yaml
Type: OfficeTextDecorationStyle
Parameter Sets: Text
Aliases: None
Possible values: None, Single, Double, Dotted, Dashed, Wavy, Words

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width for this content.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
X for this content.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Y for this content.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: True
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

- `None`

## RELATED LINKS

- None
