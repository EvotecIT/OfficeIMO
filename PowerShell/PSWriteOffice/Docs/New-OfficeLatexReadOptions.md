---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeLatexReadOptions
## SYNOPSIS
Creates native LatexParseOptions settings for document operations.

## SYNTAX
### Default (Default)
```powershell
New-OfficeLatexReadOptions [-Profile <LatexDocumentProfile>] [-MacroExpansion <LatexMacroExpansion>] [-MaximumInputLength <Int32>] [-MaximumInputBytes <Int64>] [-MaximumTokenCount <Int32>] [-MaximumNestingDepth <Int32>] [-MaximumExpansionDepth <Int32>] [-MaximumExpansionLength <Int32>] [-MaximumExpansionInputLength <Int32>] [-MaximumExpansionTokenCount <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Creates native LatexParseOptions settings for document operations.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeLatexReadOptions
```


## PARAMETERS

### -MacroExpansion
MacroExpansion override; omission preserves the native default.

```yaml
Type: LatexMacroExpansion
Parameter Sets: Default
Aliases: None
Possible values: None, SafeSimpleDefinitions

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumExpansionDepth
MaximumExpansionDepth override; omission preserves the native default.

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

### -MaximumExpansionInputLength
MaximumExpansionInputLength override; omission preserves the native default.

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

### -MaximumExpansionLength
MaximumExpansionLength override; omission preserves the native default.

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

### -MaximumExpansionTokenCount
MaximumExpansionTokenCount override; omission preserves the native default.

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

### -MaximumInputBytes
MaximumInputBytes override; omission preserves the native default.

```yaml
Type: Int64
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

### -MaximumNestingDepth
MaximumNestingDepth override; omission preserves the native default.

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

### -MaximumTokenCount
MaximumTokenCount override; omission preserves the native default.

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
Type: LatexDocumentProfile
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

- `OfficeIMO.Latex.LatexParseOptions`

## RELATED LINKS

- None
