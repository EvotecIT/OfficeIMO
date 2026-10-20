---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeProjectReport
## SYNOPSIS
Exports a captured Project report through the shared PDF, HTML, Word, Excel or PowerPoint adapters.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficeProjectReport [-Path] <string> -View <ProjectView> [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Requires PowerShell 7.4. The destination extension selects the report adapter. Rendering and editable document composition remain in OfficeIMO.Workflows.

## EXAMPLES

### EXAMPLE 1
```powershell
$view = $project | Get-OfficeProjectView -Kind Gantt
$view | Export-OfficeProjectReport -Path ./delivery.pdf
$view | Export-OfficeProjectReport -Path ./delivery.docx
```


## PARAMETERS

### -Force
Replace an existing destination.

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
Return the saved report file.

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
PDF, HTML, DOCX, XLSX or PPTX destination filename.

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

### -View
Captured report from Get-OfficeProjectView.

```yaml
Type: ProjectView
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Project.ProjectView`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
