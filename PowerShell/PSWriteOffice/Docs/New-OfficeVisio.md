---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeVisio
## SYNOPSIS
Creates a Visio document with an initial page, returning it live or saving a completed content block.

## SYNTAX
### Live (Default)
```powershell
New-OfficeVisio [[-Path] <string>] [-PageName <string>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-DisplayUnit <VisioMeasurementUnit>] [-Title <string>] [-Author <string>] [-RequestRecalcOnOpen] [-UseMastersByDefault] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficeVisio [-Path] <string> [-Content] <scriptblock> [-PageName <string>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-DisplayUnit <VisioMeasurementUnit>] [-Title <string>] [-Author <string>] [-RequestRecalcOnOpen] [-UseMastersByDefault] [-Force] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Without a block no file is written. A block publishes only after success; an existing destination requires Force. Geometry carries its own unit; DisplayUnit controls the page rulers.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeVisio -Path ./ServiceMap.vsdx -Title 'Service map' {
    VisioRectangle -Key web -Text 'Web' -X 1in -Y 4in
    VisioRectangle -Key api -Text 'API' -X 4in -Y 4in
    VisioConnector -From web -To api -Label 'calls'
}
```


### EXAMPLE 2
```powershell
$diagram = New-OfficeVisio -Width 11in -Height 8.5in
Save-OfficeVisio -Document $diagram -Path ./Landscape.vsdx
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

### -Author
Optional document author.

```yaml
Type: String
Parameter Sets: Live, Content
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

### -DisplayUnit
Measurement unit displayed by Visio, independent of geometry input units.

```yaml
Type: VisioMeasurementUnit
Parameter Sets: Live, Content
Aliases: None
Possible values: Inches, Centimeters, Millimeters

Required: False
Position: named
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

### -Height
Initial page height; defaults to A4 portrait.

```yaml
Type: OfficeLength
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageName
Initial page name.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

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

### -RequestRecalcOnOpen
Ask Visio to recalculate layout and connector routing on opening.

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

### -Title
Optional document title.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseMastersByDefault
Use Visio masters for supported built-in stencil shapes when saving.

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

### -Width
Initial page width; defaults to A4 portrait.

```yaml
Type: OfficeLength
Parameter Sets: Live, Content
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

- `OfficeIMO.Visio.VisioDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
