---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfFlatAnnotation
## SYNOPSIS
Flattens supported visual PDF annotations into static page content.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfFlatAnnotation [-Path] <string> [-OutputPath] <string> [-Password <string>] [-IgnorePermissionRestrictions] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Flattens supported visual PDF annotations into static page content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficePdfFlatAnnotation -Path .\Reviewed.pdf -OutputPath .\Reviewed-Flat.pdf
```

Paints supported annotation appearance streams into page content and writes a new PDF.

## PARAMETERS

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed modification restrictions.

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
