---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeOneNoteReadOptions
## SYNOPSIS
Creates native OneNoteLoadOptions settings for document operations.

## SYNTAX
### Default (Default)
```powershell
New-OfficeOneNoteReadOptions [-MaxInputBytes <Int64>] [-MaxObjects <Int32>] [-MaxAssetBytes <Int64>] [-MaxTotalAssetBytes <Int64>] [-NoStrictHeaderValidation] [-NoChecksumValidation] [-NoUnknownData] [<CommonParameters>]
```

## DESCRIPTION
Creates native OneNoteLoadOptions settings for document operations.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeOneNoteReadOptions
```


## PARAMETERS

### -MaxAssetBytes
MaxAssetBytes override; omission preserves the native default.

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

### -MaxInputBytes
MaxInputBytes override; omission preserves the native default.

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

### -MaxObjects
MaxObjects override; omission preserves the native default.

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

### -MaxTotalAssetBytes
MaxTotalAssetBytes override; omission preserves the native default.

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

### -NoChecksumValidation
NoChecksumValidation override; omission preserves the native default.

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

### -NoStrictHeaderValidation
NoStrictHeaderValidation override; omission preserves the native default.

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

### -NoUnknownData
NoUnknownData override; omission preserves the native default.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.OneNote.OneNoteLoadOptions`

## RELATED LINKS

- None
