---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfWord
## SYNOPSIS
Reconstructs editable Word content from a PDF.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfWord [-Path] <string> [-OutputPath] <string> [-Password <string>] [-IgnorePermissionRestrictions] [-Options <PdfToWordOptions>] [-Force] [-Open] [-PassThruReport] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Uses OfficeIMO's semantic PDF reader to recover supported headings, paragraphs, lists, tables, links, images, metadata, and page breaks.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficePdfWord -Path .\Report.pdf -OutputPath .\Report.docx
```

Writes a DOCX document and returns its file information.

## PARAMETERS

### -Force
Overwrite an existing output file.

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

### -IgnorePermissionRestrictions
After successful authentication, explicitly ignore owner-imposed extraction restrictions.

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

### -Open
Open the converted document after saving.

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
Advanced OfficeIMO PDF-to-Word reconstruction options.

```yaml
Type: PdfToWordOptions
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
Output DOCX path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: OutPath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThruReport
Return the detailed conversion report instead of file information.

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
- `OfficeIMO.Word.Pdf.PdfWordConversionReport`

## RELATED LINKS

- None
