---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePowerPointHtml
## SYNOPSIS
Converts a PowerPoint deck to an HTML review document.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficePowerPointHtml [-Path] <string> [-Password <string>] [-OutputPath <string>] [-Profile <OfficePowerPointHtmlProfile>] [-Theme <OfficeVisualThemeKind>] [-Title <string>] [-IncludeHiddenSlides] [-NoNotes] [-NoTables] [-IncludeHiddenShapes] [-NoExtractionProof] [-NoDefaultStyles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Presentation
```powershell
ConvertTo-OfficePowerPointHtml -Presentation <PowerPointPresentation> [-OutputPath <string>] [-Profile <OfficePowerPointHtmlProfile>] [-Theme <OfficeVisualThemeKind>] [-Title <string>] [-IncludeHiddenSlides] [-NoNotes] [-NoTables] [-IncludeHiddenShapes] [-NoExtractionProof] [-NoDefaultStyles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts a PowerPoint deck to an HTML review document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficePowerPointHtml -Path .\Briefing.pptx -OutputPath .\Briefing.html -Title 'Briefing Review' -PassThru
```

Loads the deck and writes an HTML file with slide text, tables, pictures, charts, and notes where available.

### EXAMPLE 2
```powershell
PS> ConvertTo-OfficePowerPointHtml -Path .\Briefing.pptx -Profile VisualReview -OutputPath .\Briefing.visual.html
```

Uses the OfficeIMO PowerPoint visual review profile.

## PARAMETERS

### -IncludeHiddenShapes
Include hidden shapes.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHiddenSlides
Include hidden slides in the HTML review output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoDefaultStyles
Do not include OfficeIMO default CSS styles.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoExtractionProof
Do not include extraction proof metadata.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoNotes
Do not include presenter notes.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTables
Do not include table content.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional output HTML path. When omitted, HTML text is returned.

```yaml
Type: String
Parameter Sets: Path, Presentation
Aliases: OutPath
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit a FileInfo when saving to disk.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open an encrypted presentation package.

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
Path to the presentation to convert.

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

### -Presentation
Presentation instance to convert.

```yaml
Type: PowerPointPresentation
Parameter Sets: Presentation
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Profile
HTML conversion profile.

```yaml
Type: OfficePowerPointHtmlProfile
Parameter Sets: Path, Presentation
Aliases: None
Possible values: SemanticSlides, VisualReview

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in HTML document theme.

```yaml
Type: OfficeVisualThemeKind
Parameter Sets: Path, Presentation
Aliases: None
Possible values: Plain, WordLike, TechnicalDocument, GitHubLike, Compact, Report

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional HTML document title.

```yaml
Type: String
Parameter Sets: Path, Presentation
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

- `System.String`
- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `System.String`
- `System.IO.FileInfo`

## RELATED LINKS

- None
