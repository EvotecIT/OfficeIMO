---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeMarkdown
## SYNOPSIS
Creates a Markdown document using a DSL scriptblock.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeMarkdown [-Path] <string> [[-Content] <scriptblock>] [-PassThru] [-NoSave] [-WriteOptions <MarkdownWriteOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <string>] [-UnorderedListMarker <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Runs the scriptblock against a Markdown document and saves it to disk unless -NoSave is specified.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeMarkdown -Path .\README.md { MarkdownHeading -Level 1 -Text 'Report'; MarkdownTable -InputObject $data }
```

Creates a README file with a heading and table content.

### EXAMPLE 2
```powershell
PS> New-OfficeMarkdown -Path .\Report.md {
  MarkdownHeading -Level 1 -Text 'Summary'
  MarkdownTable -InputObject $summary
  MarkdownHeading -Level 2 -Text 'Details'
  MarkdownTable -InputObject $details
}
```

Creates a report with two tables separated by headings.

## PARAMETERS

### -Content
DSL scriptblock describing Markdown content.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageRenderingMode
Controls how Markdown images are serialized.

```yaml
Type: MarkdownImageRenderingMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: RichMarkdown, PortableMarkdown, Html

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineEnding
Markdown line ending: CRLF, LF, CR, or a literal line ending string.

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

### -NoSave
Skip saving after executing the DSL.

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

### -PassThru
Emit a FileInfo for chaining.

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

### -Path
Destination path for the Markdown file.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath, OutputPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnorderedListMarker
Unordered list marker: '-', '*', or '+'.

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

### -WriteOptions
Optional Markdown writer options.

```yaml
Type: MarkdownWriteOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteProfile
Friendly Markdown writer profile.

```yaml
Type: OfficeMarkdownWriteProfile
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: OfficeIMO, Portable, HtmlImage

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

- `System.IO.FileInfo`
- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None
