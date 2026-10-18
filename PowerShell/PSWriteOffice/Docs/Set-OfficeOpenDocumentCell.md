---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeOpenDocumentCell
## SYNOPSIS
Sets a typed zero-based cell value in an OpenDocument spreadsheet.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeOpenDocumentCell [-Value] <Object> -Row <long> -Column <long> [-Sheet <OdsSheet>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets a typed zero-based cell value in an OpenDocument spreadsheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Set-OfficeOpenDocumentCell -Row 0 -Column 0 -Value 'Healthy'
            Set-OfficeOpenDocumentCell -Row 0 -Column 1 -Value $true
```


## PARAMETERS

### -Column
Zero-based column index.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated cell.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
Zero-based row index.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet target. Omit inside Add-OfficeOpenDocumentSheet -Content.

```yaml
Type: OdsSheet
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Value
String, number, decimal, boolean, date, date-time offset, or time span value.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.OpenDocument.OdsSheet`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdsCell`

## RELATED LINKS

- None
