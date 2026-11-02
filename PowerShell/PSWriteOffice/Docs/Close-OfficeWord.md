---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeWord
## SYNOPSIS
Closes one or more tracked Word documents, optionally saving them.

## SYNTAX
### Current (Default)
```powershell
Close-OfficeWord [-Save] [-Path <string>] [-Show] [-AllowLoss] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Close-OfficeWord [-Document] <WordDocument> [-Save] [-Path <string>] [-Show] [-AllowLoss] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### All
```powershell
Close-OfficeWord -All [-Save] [-Show] [-AllowLoss] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Provides a cmdlet wrapper for WordDocument.Dispose/Save so scripts need not call .NET methods directly. Documents must use explicit persistence to prevent writes during disposal.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Open-OfficeWord -Path .\Report.docx; Close-OfficeWord -Document $doc
```

Disposes the loaded document instance without saving changes.

### EXAMPLE 2
```powershell
PS> Close-OfficeWord
```

Closes the current tracked document when a document handle is not passed explicitly.

### EXAMPLE 3
```powershell
PS> Close-OfficeWord -Document $doc -Save -Path .\Report-final.docx -Show
```

Saves updates to Report-final.docx, opens it, and disposes the document.

## PARAMETERS

### -All
Close all tracked documents for the current runspace.

```yaml
Type: SwitchParameter
Parameter Sets: All
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowLoss
Explicitly accept reported native-save reductions.

```yaml
Type: SwitchParameter
Parameter Sets: Current, Document, All
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Word document to close.

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

### -Password
Password used to save the document as an encrypted package.

```yaml
Type: SecureString
Parameter Sets: Current, Document, All
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional target path when saving.

```yaml
Type: String
Parameter Sets: Current, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Save
Persist changes before closing.

```yaml
Type: SwitchParameter
Parameter Sets: Current, Document, All
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the file after saving. Requires -Save or -Path.

```yaml
Type: SwitchParameter
Parameter Sets: Current, Document, All
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

- `None`

## RELATED LINKS

- None
