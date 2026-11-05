---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeWordDocumentProperty
## SYNOPSIS
Gets built-in and custom document properties.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeWordDocumentProperty [-Name <string[]>] [-BuiltIn] [-Custom] [-Document <WordDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordDocumentProperty [-Path] <string[]> [-Name <string[]>] [-BuiltIn] [-Custom] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeWordDocumentProperty -LiteralPath <string[]> [-Name <string[]>] [-BuiltIn] [-Custom] [<CommonParameters>]
```

## DESCRIPTION
Gets built-in and custom document properties.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordDocumentProperty -Name @('Name')
```


### EXAMPLE 2
```powershell
Get-OfficeWordDocumentProperty -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficeWordDocumentProperty -Path @('C:\Path')
```


## PARAMETERS

### -BuiltIn
BuiltIn for this inspection.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Custom
Custom for this inspection.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document for this inspection.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
LiteralPath for this inspection.

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

### -Name
Name for this inspection.

```yaml
Type: String[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path for this inspection.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `PSWriteOffice.Models.Word.WordDocumentPropertyInfo`: Represents a Word document property exposed to PowerShell.

## RELATED LINKS

- None
