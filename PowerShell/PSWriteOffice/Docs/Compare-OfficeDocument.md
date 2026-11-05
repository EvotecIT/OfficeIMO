---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Compare-OfficeDocument
## SYNOPSIS
Compares native document content or rendered pages and optionally publishes a review artifact.

## SYNTAX
### Path (Default)
```powershell
Compare-OfficeDocument [-ReferencePath] <string> [-DifferencePath] <string> [-Visual] [-Page <OfficePageRange>] [-OutputPath <string>] [-Options <Object[]>] [-MaximumDifferences <Int32>] [-Password <securestring>] [-DifferencePassword <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Compare-OfficeDocument [-ReferenceDocument] <Object> [-DifferenceDocument] <Object> [-Visual] [-Page <OfficePageRange>] [-OutputPath <string>] [-Options <Object[]>] [-MaximumDifferences <Int32>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Compares native document content or rendered pages and optionally publishes a review artifact.

## EXAMPLES

### EXAMPLE 1
```powershell
Compare-OfficeDocument -ReferencePath 'C:\Path' -DifferencePath 'C:\Path'
```


### EXAMPLE 2
```powershell
Compare-OfficeDocument -ReferenceDocument 'Value' -DifferenceDocument 'Value'
```


## PARAMETERS

### -DifferenceDocument
Caller-owned modified native model.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DifferencePassword
Modified file's password; omission reuses Password.

```yaml
Type: SecureString
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DifferencePath
Modified document path.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumDifferences
Maximum native findings returned; Excel also enforces this during its comparison.

```yaml
Type: Int32
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Typed native comparison settings or native property overrides.

```yaml
Type: Object[]
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional Word redline, Excel findings workbook or visual HTML gallery.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
Ordered one-based pages used by visual comparison.

```yaml
Type: OfficePageRange
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Original file's password.

```yaml
Type: SecureString
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReferenceDocument
Caller-owned original native model.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReferencePath
Original document path.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Visual
Compare native rendered pages instead of semantic content.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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

- `PSWriteOffice.Models.OfficeComparisonReport`: Common comparison summary retaining the complete native findings and conversion evidence.

## RELATED LINKS

- None
