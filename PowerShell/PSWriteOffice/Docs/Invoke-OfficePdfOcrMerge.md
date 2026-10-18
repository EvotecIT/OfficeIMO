---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Invoke-OfficePdfOcrMerge
## SYNOPSIS
Runs an external OCR provider and merges recognized words with native PDF text.

## SYNTAX
### __AllParameterSets
```powershell
Invoke-OfficePdfOcrMerge [-Path] <string> -Provider <IOcrEngine> [-Options <PdfOcrMergeOptions>] [-ReadOptions <PdfLoadOptions>] [-Password <string>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

## DESCRIPTION
Runs an external OCR provider and merges recognized words with native PDF text.

## EXAMPLES

### EXAMPLE 1
```powershell
Invoke-OfficePdfOcrMerge -Path 'C:\Path' -Provider 'Value'
```


## PARAMETERS

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed extraction restrictions.

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

### -Options
Optional page selection, DPI, confidence, overlap, and limits.

```yaml
Type: PdfOcrMergeOptions
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
Source PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Provider
Engine-neutral OCR implementation.

```yaml
Type: IOcrEngine
Parameter Sets: __AllParameterSets
Aliases: Engine
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReadOptions
Optional bounded PDF parsing settings.

```yaml
Type: PdfLoadOptions
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Pdf.Ocr.PdfOcrMergeResult`

## RELATED LINKS

- None
