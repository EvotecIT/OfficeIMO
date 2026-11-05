---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfTableCell
## SYNOPSIS
Creates a native PDF cell containing text, images or form fields.

## SYNTAX
### Text (Default)
```powershell
New-OfficePdfTableCell [[-Text] <string>] [-ColumnSpan <int>] [-RowSpan <int>] [-Color <OfficeColor>] [-BackgroundColor <OfficeColor>] [-Size <OfficeLength>] [-Bold] [-Italic] [-Underline <OfficeTextDecorationStyle>] [-Strikethrough] [-Alignment <OfficeAlignment>] [-VerticalAlignment <OfficeTextVerticalAlignment>] [-Link <string>] [-LinkBookmark <string>] [-LinkContents <string>] [-Bookmark <string>] [-NoWrap] [<CommonParameters>]
```

### Content
```powershell
New-OfficePdfTableCell [-Content] <scriptblock> [-ColumnSpan <int>] [-RowSpan <int>] [-Color <OfficeColor>] [-BackgroundColor <OfficeColor>] [-Size <OfficeLength>] [-Alignment <OfficeAlignment>] [-VerticalAlignment <OfficeTextVerticalAlignment>] [-Bookmark <string>] [-NoWrap] [<CommonParameters>]
```

## DESCRIPTION
Creates a native PDF cell containing text, images or form fields.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdfTableCell -Alignment 'Value'
```


### EXAMPLE 2
```powershell
New-OfficePdfTableCell -Content { }
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
Parameter Sets: Text, Content
Aliases: None
Possible values:

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

### -Bookmark
Bookmark for this content.

```yaml
Type: String
Parameter Sets: Text, Content
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

### -ColumnSpan
ColumnSpan for this content.

```yaml
Type: Int32
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

### -Link
Link for this content.

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

### -LinkBookmark
LinkBookmark for this content.

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

### -LinkContents
LinkContents for this content.

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

### -NoWrap
NoWrap for this content.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RowSpan
RowSpan for this content.

```yaml
Type: Int32
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

Required: False
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

### -VerticalAlignment
VerticalAlignment for this content.

```yaml
Type: OfficeTextVerticalAlignment
Parameter Sets: Text, Content
Aliases: None
Possible values: Top, Center, Bottom

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

- `PSWriteOffice.Services.Table.OfficeTableCellSpec`: Describes a logical table cell that can be rendered by multiple Office table surfaces.

## RELATED LINKS

- None
