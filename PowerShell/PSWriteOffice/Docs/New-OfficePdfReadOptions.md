---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfReadOptions
## SYNOPSIS
Creates a detached native PDF read policy and optional OCR-review settings.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfReadOptions [-Profile <PdfReadProfile>] [-Page <OfficePageRange>] [-LayoutOptions <PdfTextLayoutOptions>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-OcrMergeOptions <PdfOcrMergeOptions>] [<CommonParameters>]
```

## DESCRIPTION
Creates a detached native PDF read policy and optional OCR-review settings.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePdfReadOptions -LoadOptions ([OfficeIMO.Pdf.PdfLoadOptions]::UntrustedDefaults) -Profile Fast -Page '2-4'
Read-OfficeDocument -Path ./Report.pdf -Options $options -As Text
```


### EXAMPLE 2
```powershell
$ocr = New-OfficeOcrOptions -NoLanguageDownload
$options = New-OfficePdfReadOptions -OcrMergeOptions @{ MaxPixelsPerPage = 2000000; Dpi = 96 }
Read-OfficeDocument -Path ./Scan.pdf -Options $options -As Ocr -OcrOptions $ocr
```


## PARAMETERS

### -IgnorePermissionRestrictions
Ignore permission restrictions after successful decryption; explicit false restores enforcement.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LayoutOptions
Native layout and reading-order settings.

```yaml
Type: PdfTextLayoutOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LoadOptions
Complete base policy, including strict parsing and resource ceilings.

```yaml
Type: PdfLoadOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OcrMergeOptions
Native OCR-review settings, used by Read-OfficeDocument -As Ocr.

```yaml
Type: PdfOcrMergeOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
Caller-ordered one-based source pages.

```yaml
Type: OfficePageRange
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open encrypted PDFs; omission preserves the base policy.

```yaml
Type: SecureString
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Profile
Semantic reconstruction profile; omission preserves the base policy.

```yaml
Type: PdfReadProfile
Parameter Sets: Default
Aliases: None
Possible values: Fast, Structured

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

- `PSWriteOffice.Models.Pdf.PdfReadOptions`: OCR settings apply to Read-OfficeDocument -As Ocr. The native policies remain available for direct engine calls.

## RELATED LINKS

- None
