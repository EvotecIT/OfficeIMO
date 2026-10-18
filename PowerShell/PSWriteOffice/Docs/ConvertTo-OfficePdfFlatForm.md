---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfFlatForm
## SYNOPSIS
Converts a PDF with simple AcroForm fields into a flat PDF.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfFlatForm [-Path] <string> [-OutputPath] <string> [-Password <string>] [-IgnorePermissionRestrictions] [-AppearanceFontPath <string>] [-AppearanceFontFamilyName <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts a PDF with simple AcroForm fields into a flat PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Set-OfficePdfForm -Path .\Examples\Documents\Request.pdf -OutputPath .\Examples\Documents\Request-Filled.pdf -Field @{
    Requester = 'Ada Lovelace'
    Priority = 'High'
}
ConvertTo-OfficePdfFlatForm -Path .\Examples\Documents\Request-Filled.pdf -OutputPath .\Examples\Documents\Request-Flat.pdf
```

Turns simple form fields into static page content.

## PARAMETERS

### -AppearanceFontFamilyName
PDF font family name used for the supplied appearance font.

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

### -AppearanceFontPath
TrueType or OpenType/CFF font file used to synthesize Unicode form field appearances while flattening.

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

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed form-modification restrictions.

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

### -OutputPath
Output PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to authenticate an encrypted PDF.

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
Input PDF path.

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

- `System.IO.FileInfo`

## RELATED LINKS

- None
