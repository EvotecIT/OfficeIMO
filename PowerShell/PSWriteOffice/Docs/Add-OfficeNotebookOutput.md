---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeNotebookOutput
## SYNOPSIS
Attaches saved output or supplied console text to a code cell.

## SYNTAX
### Stream (Default)
```powershell
Add-OfficeNotebookOutput -Cell <NotebookCell> -Text <string> [-Stream <string>] [-PassThru] [<CommonParameters>]
```

### Saved
```powershell
Add-OfficeNotebookOutput -Cell <NotebookCell> -Output <NotebookOutput> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Attaches saved output or supplied console text to a code cell.

## EXAMPLES

### EXAMPLE 1
```powershell
$cell | Add-OfficeNotebookOutput -Stream stdout -Text 'Lab result'
```


## PARAMETERS

### -Cell
Code cell to extend.

```yaml
Type: NotebookCell
Parameter Sets: Stream, Saved
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Output
Existing saved output to copy to the cell.

```yaml
Type: NotebookOutput
Parameter Sets: Saved
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the updated cell.

```yaml
Type: SwitchParameter
Parameter Sets: Stream, Saved
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Stream
Console stream associated with the supplied text.

```yaml
Type: String
Parameter Sets: Stream
Aliases: None
Possible values: stdout, stderr

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Saved text for the stream output.

```yaml
Type: String
Parameter Sets: Stream
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Notebook.NotebookCell`

## OUTPUTS

- `OfficeIMO.Notebook.NotebookCell`

## RELATED LINKS

- None
