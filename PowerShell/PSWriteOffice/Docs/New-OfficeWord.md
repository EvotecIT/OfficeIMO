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
### __AllParameterSets
```powershell
New-OfficeWord [[-Path] <string>] [[-Content] <scriptblock>] [-Theme <OfficeThemeInput>] [-Template <string>] [-Force] [-Password <securestring>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
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
Gets or sets whether native-save loss is explicitly accepted.

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

### -Content
Gets or sets a content block; omission returns the live document.

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

### -Force
Allow a completed content block to replace an existing destination.

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
Gets or sets whether a completed content block returns the saved FileInfo.

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

### -Password
Password for a completed content block's saved file; use Save-OfficeWord for a live document.

```yaml
Type: SecureString
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
Optional associated destination; required with a content block.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Gets or sets whether the saved file opens in the system default application.

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

### -Template
Existing Word package copied as starting content. Its styling is preserved unless Theme is supplied.

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

### -Theme
Built-in theme, theme object or a document/template whose visual theme is read.

```yaml
Type: OfficeThemeInput
Parameter Sets: __AllParameterSets
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
