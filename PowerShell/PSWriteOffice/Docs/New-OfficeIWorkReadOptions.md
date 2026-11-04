---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeIWorkReadOptions
## SYNOPSIS
Creates native IWorkLoadOptions settings for document operations.

## SYNTAX
### Default (Default)
```powershell
New-OfficeIWorkReadOptions [-NoSourceRecords] [-MaximumPackageBytes <Int64>] [-MaximumEntryCount <Int32>] [-MaximumEntryBytes <Int64>] [-MaximumTotalEntryBytes <Int64>] [-MaximumTableRows <Int32>] [-MaximumTableColumns <Int32>] [-MaximumMaterializedCells <Int32>] [-MaximumProjectedImageBytes <Int64>] [<CommonParameters>]
```

## DESCRIPTION
Creates native IWorkLoadOptions settings for document operations.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeIWorkReadOptions
```


## PARAMETERS

### -MaximumEntryBytes
MaximumEntryBytes override; omission preserves the native default.

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

### -MaximumEntryCount
MaximumEntryCount override; omission preserves the native default.

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

### -MaximumMaterializedCells
MaximumMaterializedCells override; omission preserves the native default.

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

### -MaximumPackageBytes
MaximumPackageBytes override; omission preserves the native default.

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

### -MaximumProjectedImageBytes
MaximumProjectedImageBytes override; omission preserves the native default.

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

### -MaximumTableColumns
MaximumTableColumns override; omission preserves the native default.

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

### -MaximumTableRows
MaximumTableRows override; omission preserves the native default.

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

### -MaximumTotalEntryBytes
MaximumTotalEntryBytes override; omission preserves the native default.

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

### -NoSourceRecords
NoSourceRecords override; omission preserves the native default.

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

- `OfficeIMO.IWork.IWorkLoadOptions`

## RELATED LINKS

- None
