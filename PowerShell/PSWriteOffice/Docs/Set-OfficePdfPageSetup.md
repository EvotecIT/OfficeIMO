---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePdfPageSetup
## SYNOPSIS
Sets physical dimensions, orientation and margins in the enclosing PDF composition.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfPageSetup [-PageSize <OfficePageSize>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Orientation <OfficePageOrientation>] [-Margin <OfficeLength>] [-MarginLeft <OfficeLength>] [-MarginTop <OfficeLength>] [-MarginRight <OfficeLength>] [-MarginBottom <OfficeLength>] [<CommonParameters>]
```

## DESCRIPTION
Sets physical dimensions, orientation and margins in the enclosing PDF composition.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePdfPageSetup -Height 'Value'
```


## PARAMETERS

### -Height
Custom page height.

```yaml
Type: OfficeLength
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margin
Uniform margin.

```yaml
Type: OfficeLength
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginBottom
Bottom margin.

```yaml
Type: OfficeLength
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginLeft
Left margin.

```yaml
Type: OfficeLength
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginRight
Right margin.

```yaml
Type: OfficeLength
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginTop
Top margin.

```yaml
Type: OfficeLength
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Orientation
Page orientation.

```yaml
Type: OfficePageOrientation
Parameter Sets: Context
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageSize
Named paper size or physical page dimensions.

```yaml
Type: OfficePageSize
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Custom page width.

```yaml
Type: OfficeLength
Parameter Sets: Context
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

- `None`

## RELATED LINKS

- None
