---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Read-OfficePdf
## SYNOPSIS
Reads a PDF into OfficeIMO.Pdf's canonical structured document model.

## SYNTAX
### Path (Default)
```powershell
Read-OfficePdf [-Path] <string> [-Options <PdfReadOptions>] [-Profile <PdfReadProfile>] [-PageRange <string>] [-Password <string>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

### Document
```powershell
Read-OfficePdf -Document <PdfDocument> [-Options <PdfReadOptions>] [-Profile <PdfReadProfile>] [-PageRange <string>] [<CommonParameters>]
```

## DESCRIPTION
Reads a PDF into OfficeIMO.Pdf's canonical structured document model.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $logical = Read-OfficePdf -Path .\Examples\Documents\Report.pdf -Profile Structured
foreach ($page in $logical.Pages) {
    $page.Paragraphs | ForEach-Object { $_.Text }
    $page.Tables | Select-Object @{ Name = 'Page'; Expression = { $page.PageNumber } }, @{ Name = 'Rows'; Expression = { $_.Rows.Count } }
}
```

Returns the canonical OfficeIMO.Pdf result, including typed pages, paragraphs, tables, images, and diagnostics.

## PARAMETERS

### -Document
An existing OfficeIMO.Pdf document, such as output from Get-OfficePdf.

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
After successful password authentication, explicitly ignore owner-imposed usage restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Advanced structured-read settings. Friendly parameters override the corresponding setting when explicitly supplied.

```yaml
Type: PdfReadOptions
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageRange
Optional page ranges such as 1-3,5.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open a Standard password-encrypted PDF.

```yaml
Type: String
Parameter Sets: Path
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
Parameter Sets: Path
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Profile
Semantic reconstruction profile. Structured is the default; Fast omits optional document-wide enrichment.

```yaml
Type: PdfReadProfile
Parameter Sets: Path, Document
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

- `System.String`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocumentReadResult`

## RELATED LINKS

- None
