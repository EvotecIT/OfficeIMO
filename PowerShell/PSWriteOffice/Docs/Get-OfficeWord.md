---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWord
## SYNOPSIS
Opens an existing Word document.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeWord [-Path] <string> [[-Content] <scriptblock>] [-ReadOnly] [-Password <string>] [<CommonParameters>]
```

## DESCRIPTION
Returns an OfficeIMO WordDocument for inspection, advanced operations, or optional DSL edits.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx -ReadOnly
```

Loads Report.docx and exposes the document object for querying.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx { WordParagraph -Text 'Appended by DSL' }; $doc | Save-OfficeWord
```

Loads the document, appends content through the DSL, and returns the open document for saving or further edits.

## PARAMETERS

### -Content
Optional DSL scriptblock to execute against the loaded document.

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

### -Password
Password used to open an encrypted document package.

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
Path to the .docx. Accepts PS paths.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReadOnly
Open in read-only mode.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
