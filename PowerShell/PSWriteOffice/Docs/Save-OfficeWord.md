---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeWord
## SYNOPSIS
Saves a live Word document without closing it.

## SYNTAX
### Document (Default)
```powershell
Save-OfficeWord [-Document] <WordDocument> [-Path <string>] [-Password <securestring>] [-AllowLoss] [-Show] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Uses the supplied path or associated destination. Omitted Password preserves loaded encryption. PassThru returns the saved FileInfo. The document must use explicit persistence.

## EXAMPLES

### EXAMPLE 1
```powershell
Save-OfficeWord -Document $document -Path ./Report.docx -PassThru
            Close-OfficeWord -Document $document
```


### EXAMPLE 2
```powershell
$document = Open-OfficeWord -Path ./Protected.docx -Password (Read-Host -AsSecureString)
$document | Save-OfficeWord
$document | Close-OfficeWord
```


## PARAMETERS

### -AllowLoss
Explicitly accept reported native-save reductions.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live Word document to save.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Override the output password; omission preserves the document's loaded encryption.

```yaml
Type: SecureString
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional new destination; omission uses the associated path.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved file in its registered default application.

```yaml
Type: SwitchParameter
Parameter Sets: Document
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

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
