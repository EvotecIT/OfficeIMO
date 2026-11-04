---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePdfFooter
## SYNOPSIS
Composes running PDF Footer text, child runs or the native builder inside New-OfficePdf.

## SYNTAX
### Text (Default)
```powershell
Set-OfficePdfFooter [[-Text] <string>] [-Alignment <OfficeAlignment>] [-Size <OfficeLength>] [-FirstPage] [-EvenPage] [<CommonParameters>]
```

### Content
```powershell
Set-OfficePdfFooter [-Content] <scriptblock> [-Alignment <OfficeAlignment>] [-Size <OfficeLength>] [-FirstPage] [-EvenPage] [<CommonParameters>]
```

### Builder
```powershell
Set-OfficePdfFooter -Compose <scriptblock> [<CommonParameters>]
```

## DESCRIPTION
Composes running PDF Footer text, child runs or the native builder inside New-OfficePdf.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdf ./Report.pdf { PdfFooter { PdfText 'Internal' -Italic }; PdfParagraph 'Report' }
```


## PARAMETERS

### -Alignment
Text alignment.

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

### -Compose
Advanced native Footer builder, passed as the script argument.

```yaml
Type: ScriptBlock
Parameter Sets: Builder
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Styled PdfText child commands.

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

### -EvenPage
Configure the even page variant.

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

### -FirstPage
Configure the first page variant.

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

### -Size
Text size; bare numbers mean points.

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

### -Text
Running text supporting {page} and {pages} tokens.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `System.Void`

## RELATED LINKS

- None
