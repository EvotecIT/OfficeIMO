---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeAccessColumn
## SYNOPSIS
Adds a typed field before rows are appended.

## SYNTAX
### Table (Default)
```powershell
Add-OfficeAccessColumn [-Name] <string> -Table <AccessTable> [-Type <AccessDataType>] [-MaxLength <Int32>] [-AutoNumberSeed <int>] [-Precision <int>] [-Scale <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a typed field before rows are appended.

## EXAMPLES

### EXAMPLE 1
```powershell
$table | Add-OfficeAccessColumn -Name DisplayName -Type ShortText -MaxLength 120
```


## PARAMETERS

### -AutoNumberSeed
First sequential AutoNumber value.

```yaml
Type: Int32
Parameter Sets: Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxLength
Maximum short-text length.

```yaml
Type: Int32
Parameter Sets: Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Field name.

```yaml
Type: String
Parameter Sets: Table
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the created or changed object.

```yaml
Type: SwitchParameter
Parameter Sets: Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Precision
Decimal precision.

```yaml
Type: Int32
Parameter Sets: Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Decimal scale.

```yaml
Type: Int32
Parameter Sets: Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Table in a newly created database model.

```yaml
Type: AccessTable
Parameter Sets: Table
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Type
Native field type.

```yaml
Type: AccessDataType
Parameter Sets: Table
Aliases: None
Possible values: AutoNumber, Int32, ShortText, LongText, Currency, Double, DateTime, Boolean, Guid, Binary, Byte, Int16, Single, Decimal, Int64, Complex, ExtendedDateTime, Unknown

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Access.AccessTable`

## OUTPUTS

- `OfficeIMO.Access.AccessColumn`

## RELATED LINKS

- None
