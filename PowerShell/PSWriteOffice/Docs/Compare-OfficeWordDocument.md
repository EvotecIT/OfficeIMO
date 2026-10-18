---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Compare-OfficeWordDocument
## SYNOPSIS
Compares two Word documents and optionally writes a tracked-change redline.

## SYNTAX
### __AllParameterSets
```powershell
Compare-OfficeWordDocument [-ReferencePath] <string> [-DifferencePath] <string> [-RedlinePath <string>] [-Options <WordComparisonOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Compares two Word documents and optionally writes a tracked-change redline.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $result = Compare-OfficeWordDocument -ReferencePath .\Before.docx -DifferencePath .\After.docx -RedlinePath .\Redline.docx
```

Returns deterministic findings and saves a Word document containing revision marks.

### EXAMPLE 2
```powershell
PS> $options = New-OfficeWordComparisonOptions -IgnoreWhitespace -CompareVolatileMetadata:$false
Compare-OfficeWordDocument -ReferencePath .\Before.docx -DifferencePath .\After.docx -Options $options
```


## PARAMETERS

### -DifferencePath
Path to the modified Word document.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional structural comparison switches.

```yaml
Type: WordComparisonOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RedlinePath
Optional path for a tracked-change redline document.

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

### -ReferencePath
Path to the original Word document.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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

- `None`

## OUTPUTS

- `OfficeIMO.Word.WordComparisonResult`

## RELATED LINKS

- None
