---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Test-OfficeAccessibility
## SYNOPSIS
Checks Word, Excel and PowerPoint files for accessibility issues. The input is never changed.

## SYNTAX
### Path (Default)
```powershell
Test-OfficeAccessibility [-Path] <string[]> [-Strict] [-Sarif] [-Quiet] [-IncludeWarnings] [-Password <securestring>] [<CommonParameters>]
```

### LiteralPath
```powershell
Test-OfficeAccessibility -LiteralPath <string[]> [-Strict] [-Sarif] [-Quiet] [-IncludeWarnings] [-Password <securestring>] [<CommonParameters>]
```

### Document
```powershell
Test-OfficeAccessibility -Document <Object> [-Strict] [-Sarif] [-Quiet] [-IncludeWarnings] [<CommonParameters>]
```

## DESCRIPTION
Runs the rules of the Office Accessibility Checker: alternative text, table header rows, document title and language,
headings, hyperlink text, text contrast, slide titles, merged cells, default and hidden sheets, and blank cells or paragraphs
used for spacing. All three formats return the same finding codes and severities. Strict turns structural warnings into errors
and requires a document title; heuristic findings stay warnings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> (Test-OfficeAccessibility -Path .\Budget.xlsx).Findings | Format-Table Severity, Code, Location, Message
```


### EXAMPLE 2
```powershell
PS> Get-ChildItem .\Docs -Include *.docx, *.xlsx, *.pptx -Recurse | Test-OfficeAccessibility -Strict |
                Where-Object { -not $_.IsSuccessful }
```


### EXAMPLE 3
```powershell
PS> Test-OfficeAccessibility -Path .\Deck.pptx -Sarif | Set-Content .\Deck.sarif
```


## PARAMETERS

### -Document
An open Word, Excel or PowerPoint document, including unsaved edits.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeWarnings
With Quiet, warnings also make the result false.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Exact file paths, including FileInfo.FullName and provider PSPath pipeline properties.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Password
Password for an encrypted path input. Accepts SecureString or PSCredential; plain strings are transformed with a warning.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Files to inspect; supports wildcards.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Quiet
Returns a Boolean per input: true when there are no errors (and no warnings with IncludeWarnings).

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sarif
Returns each report as a SARIF 2.1.0 log string instead of a report object.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strict
Uses the strict profile: structural warnings become errors and a document title is required.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
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

- `System.String[]`

## OUTPUTS

- `PSWriteOffice.Models.OfficeAccessibilityFileReport`: Accessibility findings for one Word, Excel or PowerPoint file or open document.
- `System.String`
- `System.Boolean`

## RELATED LINKS

- None
