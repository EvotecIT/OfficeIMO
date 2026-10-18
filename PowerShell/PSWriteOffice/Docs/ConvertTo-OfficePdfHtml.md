---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfHtml
## SYNOPSIS
Converts a PDF file to HTML through the first-party OfficeIMO HTML/PDF adapter.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfHtml [-Path] <string> [-PageRange <string>] [-Password <string>] [-IgnorePermissionRestrictions] [-OutputPath <string>] [-Profile <PdfHtmlProfile>] [-ImageExportMode <PdfHtmlImageExportMode>] [-MaxEmbeddedImageBytes <Int64>] [-NoMetadata] [-NoPageContainers] [-NoImagePlaceholders] [-IncludeLinkAnnotations] [-IncludeFormWidgets] [-Fragment] [-DocumentTitleFallback <string>] [-Options <PdfToHtmlOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts a PDF file to HTML through the first-party OfficeIMO HTML/PDF adapter.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\report.pdf { Add-OfficePdfParagraph -Text 'Ready' }
            ConvertTo-OfficePdfHtml -Path .\report.pdf -OutputPath .\report.html
```

Writes HTML generated from the OfficeIMO logical PDF read model.

## PARAMETERS

### -DocumentTitleFallback
Fallback HTML document title when PDF metadata does not provide one.

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

### -Fragment
Emit an HTML fragment instead of a complete document shell.

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

### -ImageExportMode
Controls whether extracted images are embedded or represented as placeholders.

```yaml
Type: PdfHtmlImageExportMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: PlaceholderOnly, EmbeddedDataUri

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeFormWidgets
Include AcroForm widget placeholders.

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

### -IncludeLinkAnnotations
Include link annotation placeholders.

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

### -MaxEmbeddedImageBytes
Maximum extracted image byte length that may be embedded into generated HTML. Set to 0 to disable embedding.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoImagePlaceholders
Do not emit image placeholders.

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

### -NoMetadata
Do not emit PDF metadata into the generated HTML.

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

### -NoPageContainers
Do not emit page wrapper elements.

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
Optional OfficeIMO PDF to HTML save options.

```yaml
Type: PdfToHtmlOptions
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
Optional output HTML file path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: OutPath
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
Password used to read a Standard password-encrypted PDF.

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
PDF file path.

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

### -Profile
PDF to HTML profile to use when Options is not supplied.

```yaml
Type: PdfHtmlProfile
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Semantic, PositionedReview

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

## OUTPUTS

- `System.String`
- `System.IO.FileInfo`

## RELATED LINKS

- None
