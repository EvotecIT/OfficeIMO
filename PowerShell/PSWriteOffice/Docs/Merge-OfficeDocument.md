---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Merge-OfficeDocument
## SYNOPSIS
Combines documents of the same native format in input order.

## SYNTAX
### Path (Default)
```powershell
Merge-OfficeDocument [-Path] <string[]> [-OutputPath] <string> [-Password <securestring[]>] [-Options <Object[]>] [-NoClobber] [-PassThru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Merge-OfficeDocument [-Document] <IDisposable> [-Path] <string[]> [-OutputPath <string>] [-Password <securestring[]>] [-Options <Object[]>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Merge-OfficeDocument [-OutputPath] <string> -LiteralPath <string[]> [-Password <securestring[]>] [-Options <Object[]>] [-NoClobber] [-PassThru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Supports PDF and Open XML Word and Excel files. Binary projection and conversion are not performed.
Word and Excel live documents can receive appended files and remain caller-owned. PDF merging creates a new
document, so live PDF targets are not supported. PassThru returns a saved FileInfo for file input or the live
document for Document input. Native merge options determine which content is imported.

## EXAMPLES

### EXAMPLE 1
```powershell
Merge-OfficeDocument -Path ./agenda.pdf, ./minutes.pdf -OutputPath ./board-pack.pdf -PassThru
```


### EXAMPLE 2
```powershell
Get-ChildItem ./weekly -Filter *.docx | Sort-Object Name | Merge-OfficeDocument -OutputPath ./monthly.docx
```


### EXAMPLE 3
```powershell
Merge-OfficeDocument -Document $proposal -Path ./scope.docx, ./pricing.docx -OutputPath ./proposal.docx
```


## PARAMETERS

### -Document
A single editable Word or Excel document to append to; ownership stays with the caller.

```yaml
Type: IDisposable
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Literal files in merge order, including FileInfo.FullName and provider PSPath properties.

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

### -NoClobber
Refuse to replace an existing destination; requires OutputPath.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Native merge settings, New-OfficeExcelMergeOptions results, or property hashtables for the selected format.

```yaml
Type: Object[]
Parameter Sets: Path, Document, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Destination with the same extension as every input; optional for a live document.

```yaml
Type: String
Parameter Sets: Path, Document, LiteralPath
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the saved FileInfo, or the caller-owned live document for Document input.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
One source password reused for all files, or one password per resolved file in order.

```yaml
Type: SecureString[]
Parameter Sets: Path, Document, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Files in merge order, or files appended to Document; wildcard matches retain provider order.

```yaml
Type: String[]
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Show
Open the saved result through the shell after successful publication.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `System.IDisposable`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Excel.ExcelDocument`

## RELATED LINKS

- None
