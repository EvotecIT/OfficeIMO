---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeAsciiDocReadOptions
## SYNOPSIS
Creates native AsciiDocParseOptions settings for document operations.

## SYNTAX
### Default (Default)
```powershell
New-OfficeAsciiDocReadOptions [-Profile <AsciiDocDocumentProfile>] [-MaximumInputLength <Int32>] [-MaximumBlockCount <Int32>] [-MaximumInlineNestingDepth <Int32>] [-MaximumInlineNodeCount <Int32>] [-MaximumTableCellCount <Int32>] [-MaximumTableColumnCount <Int32>] [-MaximumTableSpan <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Creates native AsciiDocParseOptions settings for document operations.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeAsciiDocReadOptions
```


## PARAMETERS

### -MaximumBlockCount
MaximumBlockCount override; omission preserves the native default.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInlineNestingDepth
MaximumInlineNestingDepth override; omission preserves the native default.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInlineNodeCount
MaximumInlineNodeCount override; omission preserves the native default.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInputLength
MaximumInputLength override; omission preserves the native default.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTableCellCount
MaximumTableCellCount override; omission preserves the native default.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTableColumnCount
MaximumTableColumnCount override; omission preserves the native default.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTableSpan
MaximumTableSpan override; omission preserves the native default.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Profile
Profile override; omission preserves the native default.

```yaml
Type: AsciiDocDocumentProfile
Parameter Sets: Default
Aliases: None
Possible values: OfficeIMO, PreserveOnly

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

- `OfficeIMO.AsciiDoc.AsciiDocParseOptions`

## RELATED LINKS

- None
