---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfCanvas
## SYNOPSIS
Draws arbitrary visual canvas content on existing PDF pages.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePdfCanvas [-Content] <scriptblock> -Path <string> -OutputPath <string> [-PageRange <string>] [-BehindContent] [-Opacity <double>] [-ConfigureRendering <scriptblock>] [-Password <string>] [-IgnorePermissionRestrictions] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The script receives a PdfPageCanvas and PdfStampPageContext.
It can draw text, rich text, images, shapes, drawings, and tables. Interactive annotations,
links, form fields, and outlines are separate PDF operations and are rejected by this visual-only surface.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficePdfCanvas -Path .\Report.pdf -OutputPath .\Stamped.pdf -PageRange '1,last' -Content {
    param($canvas, $page)
    $null = $canvas.Text("Review copy $($page.PageNumber)/$($page.PageCount)", 36, 24, $page.Width - 72, 24, 10)
}
```

The callback runs once for every selected page and may mix any supported visual canvas primitives.

## PARAMETERS

### -BehindContent
Place the generated canvas behind existing page content.

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

### -ConfigureRendering
The callback receives a PdfOptions instance. Page geometry, margins, and encryption remain controlled by the stamping operation.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Canvas callback. Declare parameters for the canvas and page context.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed usage restrictions.
This does not discover, bypass, or crack a missing password.

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

### -Opacity
Opacity applied to the complete generated canvas.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Output PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageRange
Target page selector such as 1-3,odd,last. Omit to stamp every page.

```yaml
Type: String
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
Emit the object created or changed by the command.

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
Password used to authenticate an encrypted input PDF.

```yaml
Type: String
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
Input PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
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

- `None`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
