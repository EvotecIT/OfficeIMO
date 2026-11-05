---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeWord
## SYNOPSIS
Opens a Word document for explicit editing, or saves and closes a completed content block.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeWord [-Path] <string> [[-Content] <scriptblock>] [-PassThru] [-AllowLoss] [-Show] [-ReadOnly] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeWord [[-Content] <scriptblock>] -LiteralPath <string> [-PassThru] [-AllowLoss] [-Show] [-ReadOnly] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Opening preserves the document's theme and encryption. Without a block the document remains live. A failed block leaves the source unchanged; read-only input cannot run an authoring block.

## EXAMPLES

### EXAMPLE 1
```powershell
$document = Get-Item './Review[final].docx' | Open-OfficeWord
$document | Close-OfficeWord
```


### EXAMPLE 2
```powershell
Open-OfficeWord -Path ./Report.docx -PassThru {
    WordParagraph 'Reviewed and approved'
}
```


## PARAMETERS

### -AllowLoss
Explicitly accept native-save loss for an authoring block.

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

### -Content
An optional authoring block; omission returns the live document.

```yaml
Type: ScriptBlock
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Word document path without wildcard expansion; accepts FileInfo.FullName and provider PSPath.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -PassThru
Emit the FileInfo after a completed authoring block.

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

### -Password
Source password as a SecureString, credential, or warned plaintext value.

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
Word document path with FileSystem wildcard expansion.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -ReadOnly
Return a read-only live document. Cannot be combined with Content.

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

### -Show
Open the saved document after a completed authoring block.

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

- `System.String`

## OUTPUTS

- `OfficeIMO.Word.WordDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
