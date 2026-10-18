---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfTheme
## SYNOPSIS
Applies an OfficeIMO.Pdf theme preset to a generated PDF document.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfTheme [-Theme] <OfficePdfThemePreset> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfTheme [-Theme] <OfficePdfThemePreset> -Document <PdfDocument> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Themes provide a reusable visual baseline for generated PDFs. They are OfficeIMO.Pdf presets, exposed by PSWriteOffice as simple enum values.
Apply a theme near the start of a New-OfficePdf script block so later content inherits the intended report rhythm.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Report.pdf {
  PdfTheme Report
  PdfHeading 'Service Review'
  PdfParagraph 'The report theme defines a polished baseline.'
}
```

Uses the PDF report theme for generated content.

## PARAMETERS

### -Document
Compatibility parameter. Page composition is supported only inside New-OfficePdf.

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

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Theme preset to apply.

```yaml
Type: OfficePdfThemePreset
Parameter Sets: Context, Document
Aliases: None
Possible values: WordLike, TechnicalDocument, Compact, Report

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
