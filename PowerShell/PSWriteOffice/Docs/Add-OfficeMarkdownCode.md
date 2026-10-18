---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownCode
## SYNOPSIS
Adds a Markdown code block.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownCode [[-Language] <string>] [-Content] <string> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownCode [[-Language] <string>] [-Content] <string> -Document <MarkdownDoc> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown code block.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $script = @(
    'Import-Module PSWriteOffice'
    'Get-OfficeExcelSummary -Path .\report.xlsx'
) -join [Environment]::NewLine
New-OfficeMarkdown -Path .\Cookbook.md {
    Add-OfficeMarkdownHeading -Level 2 -Text 'Workbook summary'
    Add-OfficeMarkdownCode -Language 'powershell' -Content $script
}
```

Appends a fenced code block with a language identifier.

## PARAMETERS

### -Content
Code content.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Markdown document to update outside the DSL context.

```yaml
Type: MarkdownDoc
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Language
Code language identifier.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the Markdown document after appending the code block.

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

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None
