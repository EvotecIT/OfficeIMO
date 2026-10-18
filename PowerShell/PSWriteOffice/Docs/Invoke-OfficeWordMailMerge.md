---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Invoke-OfficeWordMailMerge
## SYNOPSIS
Executes a simple mail merge against MERGEFIELD values in a Word document.

## SYNTAX
### __AllParameterSets
```powershell
Invoke-OfficeWordMailMerge [-InputObject] <Object> [-Document <WordDocument>] [-PreserveFields] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Executes a simple mail merge against MERGEFIELD values in a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Invoke-OfficeWordMailMerge -InputObject @{ FirstName = 'John'; OrderId = 12345 }
```

Updates MERGEFIELD values in the active Word document.

## PARAMETERS

### -Document
Document to update when provided explicitly.

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

### -InputObject
Hashtable or object whose properties map to MERGEFIELD names.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: Data, Values
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated document.

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

### -PreserveFields
Preserve field codes and only update displayed field text.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordDocument`

## RELATED LINKS

- None
