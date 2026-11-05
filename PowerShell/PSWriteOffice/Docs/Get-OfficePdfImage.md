---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePdfImage
## SYNOPSIS
Lists image resources, or extracts them with -OutputDirectory.

## SYNTAX
### Path (Default)
```powershell
Get-OfficePdfImage [-Path] <string[]> [-Page <OfficePageRange>] [-OutputDirectory <string>] [-BaseName <string>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficePdfImage -LiteralPath <string[]> [-Page <OfficePageRange>] [-OutputDirectory <string>] [-BaseName <string>] [-NoClobber] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Get-OfficePdfImage -Document <PdfDocument> [-Page <OfficePageRange>] [-OutputDirectory <string>] [-BaseName <string>] [-NoClobber] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Lists image resources, or extracts them with -OutputDirectory.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePdfImage -Path @('C:\Path')
```


### EXAMPLE 2
```powershell
Get-OfficePdfImage -Document 'Value'
```


### EXAMPLE 3
```powershell
Get-OfficePdfImage -LiteralPath @('C:\Path')
```


## PARAMETERS

### -BaseName
BaseName for PDF inspection or extraction.

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

### -Document
Document for PDF inspection or extraction.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
IgnorePermissionRestrictions for PDF inspection or extraction.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
LiteralPath for PDF inspection or extraction.

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

### -LoadOptions
LoadOptions for PDF inspection or extraction.

```yaml
Type: PdfLoadOptions
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
NoClobber for PDF inspection or extraction.

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

### -OutputDirectory
OutputDirectory for PDF inspection or extraction.

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

### -Page
Page for PDF inspection or extraction.

```yaml
Type: OfficePageRange
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
Password for PDF inspection or extraction.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path for PDF inspection or extraction.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfExtractedImage`
- `System.IO.FileInfo`

## RELATED LINKS

- None
