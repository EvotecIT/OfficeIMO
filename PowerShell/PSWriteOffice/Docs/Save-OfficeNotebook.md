---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeNotebook
## SYNOPSIS
Atomically saves a live notebook as .ipynb and leaves it open for further editing.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeNotebook [[-Path] <string>] -Document <NotebookDocument> [-AllowLoss] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Atomically saves a live notebook as .ipynb and leaves it open for further editing.

## EXAMPLES

### EXAMPLE 1
```powershell
$notebook | Save-OfficeNotebook -Path ./updated.ipynb -NoClobber -PassThru
```


## PARAMETERS

### -AllowLoss
Accept loss reported when normalizing the source to Jupyter notebook format.

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

### -Document
Live notebook to save.

```yaml
Type: NotebookDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoClobber
Refuse to replace a destination that already exists.

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

### -PassThru
Return the saved file.

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

### -Path
.ipynb destination. Omission uses the path associated with a native notebook at creation or open time.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Notebook.NotebookDocument`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
