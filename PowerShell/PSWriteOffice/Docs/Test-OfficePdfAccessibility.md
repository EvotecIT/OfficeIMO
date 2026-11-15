---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Test-OfficePdfAccessibility
## SYNOPSIS
Checks a PDF against PDF/UA-1 using the Matterhorn Protocol failure conditions.

## SYNTAX
### Path (Default)
```powershell
Test-OfficePdfAccessibility [-Path] <string[]> [-LoadOptions <PdfLoadOptions>] [-Password <securestring>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

### Document
```powershell
Test-OfficePdfAccessibility -Document <PdfDocument> [<CommonParameters>]
```

### LiteralPath
```powershell
Test-OfficePdfAccessibility -LiteralPath <string[]> [-LoadOptions <PdfLoadOptions>] [-Password <securestring>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

## DESCRIPTION
Works on any PDF, including files OfficeIMO did not produce. Machine-checkable conditions are evaluated;
conditions that need a person's judgement are returned with status NeedsReview and are never reported as passed.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $report = Test-OfficePdfAccessibility -Path .\Report.pdf
$report.FailedConditions | Select-Object @{ n = 'Condition'; e = { $_.Condition.ConditionId } }, OccurrenceCount
```

Returns the accessibility report; Findings carry page, object and structure-element locations.

### EXAMPLE 2
```powershell
PS> (Test-OfficePdfAccessibility -Path .\Report.pdf).Findings |
                Where-Object Severity -eq 'Error' |
                Format-Table ConditionId, PageNumber, ObjectNumber, StructureType, Message
```

Lists the failing objects with the Matterhorn condition they break.

## PARAMETERS

### -Document
PDF document to check.

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

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed usage restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal PDF file paths to check.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
Native input loading policy.

```yaml
Type: PdfLoadOptions
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open a Standard password-encrypted PDF.

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
PDF file paths to check; wildcards are expanded.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`
- `System.String[]`

## OUTPUTS

- `OfficeIMO.Pdf.PdfAccessibilityReport`

## RELATED LINKS

- None
