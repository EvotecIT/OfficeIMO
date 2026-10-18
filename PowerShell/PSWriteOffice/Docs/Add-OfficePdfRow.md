---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfRow
## SYNOPSIS
Adds a semantic row with percentage-based columns to a generated PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfRow [-Column] <Object[]> [-Gap <Double>] [-SpacingBefore <Double>] [-SpacingAfter <Double>] [-KeepTogether] [-KeepWithNext] [-ColumnSeparatorColor <string>] [-ColumnSeparatorWidth <Double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfRow [-Column] <Object[]> -Document <PdfDocument> [-Gap <Double>] [-SpacingBefore <Double>] [-SpacingAfter <Double>] [-KeepTogether] [-KeepWithNext] [-ColumnSeparatorColor <string>] [-ColumnSeparatorWidth <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Rows are intended for report-style layouts where two or more content groups should sit beside each other in the normal PDF flow.
Column widths are percentages and default to an even split. Column content may use headings, paragraphs, panels, lists, tables,
horizontal rules, spacers, bookmarks, or rich Run/Runs text specifications.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Report.pdf {
  PdfRow -Gap 16 -Column @(
    @{ Width = 35; Content = @(
      @{ Type = 'Heading'; Level = 2; Text = 'Signals' }
      @{ Type = 'List'; Items = @('Healthy', 'Watch', 'Needs action') }
    ) }
    @{ Width = 65; Content = @(
      @{ Type = 'Panel'; Text = 'Right-side callout content.' }
    ) }
  )
}
```

Adds a row with list content on the left and a panel on the right.

### EXAMPLE 2
```powershell
PS> New-OfficePdf -Path .\Report.pdf {
  PdfBookmark 'details'
  PdfRow -Column @(
    @{ Content = @(
      @{ Type = 'Paragraph'; Run = @(
        @{ Text = 'Jump to ' }
        @{ Text = 'details'; LinkDestinationName = 'details'; Color = '#7C3AED' }
      ) }
    ) }
  )
}
```

Uses the same rich run model as Add-OfficePdfText inside a row layout.

## PARAMETERS

### -Column
Column specifications. Each entry may define Width and Content, or shorthand values such as Heading, Paragraph, Run, Panel, List, Table, Rule, Spacer, and Bookmark.

```yaml
Type: Object[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnSeparatorColor
Optional vertical separator color between columns.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnSeparatorWidth
Optional vertical separator width in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

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

### -Gap
Horizontal gutter between columns in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepTogether
Keep the row together when possible.

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

### -KeepWithNext
Keep the row with the next visible block when possible.

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

### -SpacingAfter
Vertical spacing after the row in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingBefore
Vertical spacing before the row in PDF points.

```yaml
Type: Double
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
