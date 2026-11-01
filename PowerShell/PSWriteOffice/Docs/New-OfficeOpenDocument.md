---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeOpenDocument
## SYNOPSIS
Creates an ODT, ODS, ODP or ODG model, returning it live or saving a completed content block.

## SYNTAX
### Live (Default)
```powershell
New-OfficeOpenDocument [[-Path] <string>] [-Kind <OdfDocumentKind>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficeOpenDocument [-Path] <string> [-Content] <scriptblock> [-Kind <OdfDocumentKind>] [-Force] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Kind is inferred from Path when omitted. Without a block no file is written. A block publishes only after success; an existing destination requires Force.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeOpenDocument -Path ./Report.odt {
    Add-OfficeOpenDocumentHeading -Text 'Service report' -Level 1
    Add-OfficeOpenDocumentParagraph -Text 'All services are available.'
}
```


### EXAMPLE 2
```powershell
$drawing = New-OfficeOpenDocument -Path ./Overview.odg
$page = $drawing.AddPage('Overview')
Save-OfficeOpenDocument -Document $drawing
```


## PARAMETERS

### -AllowLoss
Explicitly accept reductions reported by the content block's native save.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Authoring block to complete before the document is atomically saved and closed.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Allow a completed content block to replace an existing destination.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Kind
Document kind; required when no destination extension is supplied.

```yaml
Type: OdfDocumentKind
Parameter Sets: Live, Content
Aliases: None
Possible values: Text, Spreadsheet, Presentation, Graphics

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo after a completed content block.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional associated destination for a live document; required when a content block saves a file.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved result in its registered application after the content block succeeds.

```yaml
Type: SwitchParameter
Parameter Sets: Content
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

- `None`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdfDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
