---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeRtfParagraph
## SYNOPSIS
Adds a paragraph to an RTF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeRtfParagraph [-Text] <string> [-Bold] [-Italic] [-Underline <RtfUnderlineStyle>] [-Strikethrough] [-Font <string>] [-Size <OfficeLength>] [-Color <OfficeColor>] [-Link <uri>] [-Alignment <OfficeAlignment>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeRtfParagraph [-Text] <string> [-Bold] [-Italic] [-Underline <RtfUnderlineStyle>] [-Strikethrough] [-Font <string>] [-Size <OfficeLength>] [-Color <OfficeColor>] [-Link <uri>] [-Alignment <OfficeAlignment>] [-PassThru] [-Document <RtfDocument>] [<CommonParameters>]
```

## DESCRIPTION
Adds a paragraph to an RTF document.

## EXAMPLES

### EXAMPLE 1
```powershell
$rtf = New-OfficeRtf
Add-OfficeRtfParagraph -Document $rtf -Text 'Maintenance completed.'
Save-OfficeRtf -Document $rtf -Path .\Notes\Maintenance.rtf
```


## PARAMETERS

### -Alignment
Alignment for this operation.

```yaml
Type: OfficeAlignment
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Bold for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document for this operation.

```yaml
Type: RtfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Font
Font family registered in the document font table.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Italic for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Link
Hyperlink target written as a native RTF field.

```yaml
Type: Uri
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
PassThru for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Apply strikethrough; explicit false clears it.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text for this operation.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Native RTF underline variant.

```yaml
Type: RtfUnderlineStyle
Parameter Sets: Context, Document
Aliases: None
Possible values: None, Single, Words, Double, Dotted, Dash, DashDot, DashDotDot, Thick, ThickDotted, ThickDash, ThickDashDot, ThickDashDotDot, Wave, HeavyWave, DoubleWave, LongDash, ThickLongDash

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Rtf.RtfDocument`

## OUTPUTS

- `OfficeIMO.Rtf.RtfDocument`

## RELATED LINKS

- None
