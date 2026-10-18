---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfBookmark
## SYNOPSIS
Adds a named bookmark at the current generated PDF flow position.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfBookmark [-Name] <string> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfBookmark [-Name] <string> -Document <PdfDocument> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a named bookmark at the current generated PDF flow position.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfBookmarks.pdf {
  Add-OfficePdfText -Run @(
    @{ Text = 'Jump to details'; LinkDestinationName = 'details'; Color = '#2563EB'; Underline = $true }
  )
  Add-OfficePdfPageBreak
  Add-OfficePdfBookmark -Name 'details'
  Add-OfficePdfHeading -Text 'Details' -Level 2
}
```

Creates an internal link target inside the generated PDF.

## PARAMETERS

### -Document
PDF document to update outside the DSL context.

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

### -Name
Bookmark name.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
