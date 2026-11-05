---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePdfFont
## SYNOPSIS
Lists fonts with embedding and ToUnicode diagnostics.

## SYNTAX
### Path (Default)
```powershell
Get-OfficePdfFont [-Path] <string[]> [-Subtype <string>] [-NeedsReview] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficePdfFont -LiteralPath <string[]> [-Subtype <string>] [-NeedsReview] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficePdfFont -Document <PdfDocument> [-Subtype <string>] [-NeedsReview] [<CommonParameters>]
```

## DESCRIPTION
Lists fonts with embedding and ToUnicode diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePdfFont -Path @('C:\Path')
```


### EXAMPLE 2
```powershell
Get-OfficePdfFont -Document 'Value'
```


### EXAMPLE 3
```powershell
Get-OfficePdfFont -LiteralPath @('C:\Path')
```


## PARAMETERS

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

### -NeedsReview
NeedsReview for PDF inspection or extraction.

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

### -Subtype
Subtype for PDF inspection or extraction.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfFontDiagnostic`

## RELATED LINKS

- None
