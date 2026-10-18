---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfMarkdown
## SYNOPSIS
Converts PDF logical text readback to Markdown.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfMarkdown [-Path] <string> [-PageRange <string>] [-Password <string>] [-IgnorePermissionRestrictions] [-OutputPath <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts PDF logical text readback to Markdown.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $markdownPath = '.\Examples\Documents\Report.md'
ConvertTo-OfficePdfMarkdown -Path .\Examples\Documents\Report.pdf -PageRange '1-3' -OutputPath $markdownPath
Get-Content $markdownPath -TotalCount 20
```

Writes Markdown readback for selected pages to a file.

## PARAMETERS

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed extraction restrictions.

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

### -OutputPath
Optional output Markdown file path.

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

### -PageRange
Optional page ranges such as 1-3,5.

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

### -Password
Password used to read a Standard password-encrypted PDF.

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
PDF file path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `System.String`
- `System.IO.FileInfo`

## RELATED LINKS

- None
