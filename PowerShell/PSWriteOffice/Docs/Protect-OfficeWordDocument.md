---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Protect-OfficeWordDocument
## SYNOPSIS
Protects a Word document with a password.

## SYNTAX
### __AllParameterSets
```powershell
Protect-OfficeWordDocument [-Password] <string> [-Document <WordDocument>] [-ProtectionType <WordDocumentProtectionType>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets the protection password and the protection type (default: ReadOnly).

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\ProtectedReport.docx {
    Add-OfficeWordParagraph -Text 'Confidential report'
    Protect-OfficeWordDocument -Password 'secret' -ProtectionType ReadOnly
}
```

Sets OfficeIMO.Word document protection while the document is being composed.

## PARAMETERS

### -Document
Document to protect when provided explicitly.

```yaml
Type: WordDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the protected document.

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

### -Password
Password to apply.

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

### -ProtectionType
Protection type (defaults to ReadOnly).

```yaml
Type: WordDocumentProtectionType
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, ReadOnly, Comments, TrackedChanges, Forms

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
