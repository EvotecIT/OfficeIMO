---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeWord
## SYNOPSIS
Creates a themed Word document, returning a live document or saving a completed content block.

## SYNTAX
### Live (Default)
```powershell
New-OfficeWord [[-Path] <string>] [-Theme <OfficeThemeInput>] [-Template <string>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficeWord [-Path] <string> [-Content] <scriptblock> [-Theme <OfficeThemeInput>] [-Template <string>] [-Force] [-Password <securestring>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Without a block no file is written, including on disposal. A block saves atomically; an existing destination requires Force. Template copies content, while Theme reads only the source theme.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeWord -Path ./Report.docx -Theme Modern -PassThru {
    WordHeading 'Service review'
    WordParagraph 'All services are available.'
}
```


### EXAMPLE 2
```powershell
$document = New-OfficeWord -Theme Boardroom
$document | Add-OfficeWordParagraph -Text 'Reviewed and approved'
Save-OfficeWord -Document $document -Path ./Review.docx
Close-OfficeWord -Document $document
```


## PARAMETERS

### -AllowLoss
Explicitly accept reductions reported by the content block's native save.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Authoring block to complete before the document is atomically saved and closed.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Allow a completed content block to replace an existing destination.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo after a completed content block.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password for a completed content block's saved file; use Save-OfficeWord for a live document.

```yaml
Type: SecureString
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional associated destination for a live document; required when a content block saves a file.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved result in its registered application after the content block succeeds.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Template
Existing Word package copied as starting content. Its styling is preserved unless Theme is supplied.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in theme, theme object or a document/template whose visual theme is read.

```yaml
Type: OfficeThemeInput
Parameter Sets: Live, Content
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

- `None`

## OUTPUTS

- `OfficeIMO.Word.WordDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
