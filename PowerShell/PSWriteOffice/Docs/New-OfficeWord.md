---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeWord
## SYNOPSIS
Creates a Word document using the DSL.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeWord [-Path] <string> [[-Content] <scriptblock>] [-TemplatePath <string>] [-PassThru] [-Open] [-NoSave] [-Password <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Handles file creation or template cloning, scriptblock execution, explicit save or live-document composition, and emits the document path when -PassThru is used.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx { WordSection { WordParagraph 'Hello DSL' } } -Open
```

Builds a document, adds one paragraph, saves it to disk, and opens it.

### EXAMPLE 2
```powershell
PS> New-OfficeWord -TemplatePath .\Template.docx -Path .\Report.docx { WordParagraph -Text 'Generated content' -StyleId 'ReportBody' }
```

Copies the template to the output path, runs the DSL against the copied document, and saves it.

### EXAMPLE 3
```powershell
PS> $document = New-OfficeWord -Path .\Report.docx -NoSave
$document | Add-OfficeWordParagraph -Text 'Status report' -Style Heading1
$document | Save-OfficeWord
$document | Close-OfficeWord
```

Associates the output path with a live document, adds content through the pipeline, then saves and closes it once.

## PARAMETERS

### -Content
DSL scriptblock describing document content.

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

### -Open
Open the document after saving.

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

### -Password
Password used to save the document as an encrypted package.

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
Destination path for the document.

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

### -TemplatePath
Existing .docx file to clone before running the DSL.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
