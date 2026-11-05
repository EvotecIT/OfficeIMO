---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Unprotect-OfficeDocument
## SYNOPSIS
Removes native package encryption or selected editing restrictions.

## SYNTAX
### Path (Default)
```powershell
Unprotect-OfficeDocument [-Path] <string[]> [-Password <securestring>] [-Restriction <OfficeProtectionRestriction[]>] [-OutputPath <string>] [-Options <Object[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Unprotect-OfficeDocument -LiteralPath <string[]> [-Password <securestring>] [-Restriction <OfficeProtectionRestriction[]>] [-OutputPath <string>] [-Options <Object[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Unprotect-OfficeDocument [-Document] <Object> [-Password <securestring>] [-Restriction <OfficeProtectionRestriction[]>] [-OutputPath <string>] [-Options <Object[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Removes native package encryption or selected editing restrictions.

## EXAMPLES

### EXAMPLE 1
```powershell
Unprotect-OfficeDocument -Path @('C:\Path')
```


### EXAMPLE 2
```powershell
Unprotect-OfficeDocument -Document 'Value'
```


### EXAMPLE 3
```powershell
Unprotect-OfficeDocument -LiteralPath @('C:\Path')
```


## PARAMETERS

### -Document
Caller-owned native model.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Exact input paths.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Native source policy options; PDF accepts its bounded load settings.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Write one copy; omission saves to the associated path when available.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the saved file or caller-owned/native replacement model.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Current open password, or PDF owner password for removing Standard security.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Input paths; supports wildcards.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Restriction
Remove only these restrictions; omission also removes package encryption.

```yaml
Type: OfficeProtectionRestriction[]
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values: ReadOnly, Comments, TrackedChanges, Forms, Structure, Windows, ReadOnlyRecommended, Print, Copy, Modify, Annotate, FillForms, Assemble

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `System.Object`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Excel.ExcelDocument`
- `OfficeIMO.PowerPoint.PowerPointDocument`
- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
