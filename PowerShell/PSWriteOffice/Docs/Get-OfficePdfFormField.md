---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfFormField
## SYNOPSIS
Gets simple AcroForm fields from a PDF.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePdfFormField [-Path] <string> [-Name <string>] [-Password <string>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

## DESCRIPTION
Gets simple AcroForm fields from a PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePdfFormField -Path .\Examples\Documents\Request.pdf |
    Select-Object Name, FieldType, Value
Set-OfficePdfForm -Path .\Examples\Documents\Request.pdf -OutputPath .\Examples\Documents\Request-Filled.pdf -Field @{ Requester = 'Ada Lovelace' }
```

Reads form field names so the fill hashtable can use the right keys.

## PARAMETERS

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed usage restrictions.

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

### -Name
Optional field name filter.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to inspect a Standard password-encrypted PDF.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
PDF file path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Pdf.PdfFormField`

## RELATED LINKS

- None
