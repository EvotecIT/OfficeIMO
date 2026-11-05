---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeWordContentControl
## SYNOPSIS
Gets typed native content controls filtered by family, alias, tag, rendered text or checkbox state.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeWordContentControl [-Document <WordDocument>] [-Type <WordContentControlType[]>] [-Alias <string[]>] [-Tag <string[]>] [-Text <string[]>] [-State <WordCheckBoxState>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordContentControl [-Path] <string[]> [-Type <WordContentControlType[]>] [-Alias <string[]>] [-Tag <string[]>] [-Text <string[]>] [-State <WordCheckBoxState>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeWordContentControl -LiteralPath <string[]> [-Type <WordContentControlType[]>] [-Alias <string[]>] [-Tag <string[]>] [-Text <string[]>] [-State <WordCheckBoxState>] [<CommonParameters>]
```

## DESCRIPTION
Gets typed native content controls filtered by family, alias, tag, rendered text or checkbox state.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordContentControl -Alias @('Value')
```


### EXAMPLE 2
```powershell
Get-OfficeWordContentControl -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficeWordContentControl -Path @('C:\Path')
```


## PARAMETERS

### -Alias
Wildcard control aliases.

```yaml
Type: String[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Caller-owned document; omission reads the enclosing Word document.

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
Literal Word files, including FileInfo.FullName.

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

### -Path
Word files with FileSystem wildcard expansion.

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

### -State
Checked or unchecked; implies the checkbox family.

```yaml
Type: WordCheckBoxState
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values: Checked, Unchecked

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Tag
Wildcard control tags.

```yaml
Type: String[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Wildcard native rendered text.

```yaml
Type: String[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Type
Control families; omission includes all seven native families.

```yaml
Type: WordContentControlType[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values: Text, CheckBox, ComboBox, DatePicker, DropDownList, Picture, RepeatingSection

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
- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordStructuredDocumentTag`
- `OfficeIMO.Word.WordCheckBox`
- `OfficeIMO.Word.WordComboBox`
- `OfficeIMO.Word.WordDatePicker`
- `OfficeIMO.Word.WordDropDownList`
- `OfficeIMO.Word.WordPictureControl`
- `OfficeIMO.Word.WordRepeatingSection`

## RELATED LINKS

- None
