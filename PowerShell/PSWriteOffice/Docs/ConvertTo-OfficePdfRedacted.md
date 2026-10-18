---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfRedacted
## SYNOPSIS
Applies explicit or text-matched PDF redactions and writes a new PDF.

## SYNTAX
### Rectangle (Default)
```powershell
ConvertTo-OfficePdfRedacted [-Path] <string> [-OutputPath] <string> -PageNumber <int> -X <double> -Y <double> -Width <double> -Height <double> [-Label <string>] [-FillColor <string>] [-OnlyPaintMatches] [-Password <string>] [-IgnorePermissionRestrictions] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Area
```powershell
ConvertTo-OfficePdfRedacted [-Path] <string> [-OutputPath] <string> -Area <PdfRedactionArea[]> [-FillColor <string>] [-OnlyPaintMatches] [-Password <string>] [-IgnorePermissionRestrictions] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
ConvertTo-OfficePdfRedacted [-Path] <string> [-OutputPath] <string> -Text <string[]> [-MatchCase] [-FillColor <string>] [-OnlyPaintMatches] [-Password <string>] [-IgnorePermissionRestrictions] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Applies explicit or text-matched PDF redactions and writes a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficePdfRedacted -Path .\Report.pdf -OutputPath .\Report-Redacted.pdf -PageNumber 1 -X 72 -Y 650 -Width 240 -Height 32
```

Removes matching text objects and annotations in the rectangle, then paints a redaction mark.

### EXAMPLE 2
```powershell
PS> ConvertTo-OfficePdfRedacted -Path .\Report.pdf -OutputPath .\Report-Redacted.pdf -Text 'Account number'
```

Finds matching logical text blocks, removes their content, and paints the derived areas.

## PARAMETERS

### -Area
One or more pre-created OfficeIMO.Pdf redaction areas.

```yaml
Type: PdfRedactionArea[]
Parameter Sets: Area
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FillColor
Redaction fill color. Named colors and hexadecimal values are accepted. Defaults to black.

```yaml
Type: String
Parameter Sets: Rectangle, Area, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Rectangle height in PDF points.

```yaml
Type: Double
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed modification restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Rectangle, Area, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Label
Optional redaction area label.

```yaml
Type: String
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MatchCase
Use case-sensitive literal text matching.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OnlyPaintMatches
Paint only areas that match text or annotations in the redaction plan.

```yaml
Type: SwitchParameter
Parameter Sets: Rectangle, Area, Text
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
Parameter Sets: Rectangle, Area, Text
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageNumber
One-based page number for the redaction rectangle.

```yaml
Type: Int32
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to read a Standard password-encrypted PDF.

```yaml
Type: String
Parameter Sets: Rectangle, Area, Text
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
Parameter Sets: Rectangle, Area, Text
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Literal text used to derive redaction areas from matching logical text blocks.

```yaml
Type: String[]
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Rectangle width in PDF points.

```yaml
Type: Double
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Left coordinate in PDF points.

```yaml
Type: Double
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Bottom coordinate in PDF points.

```yaml
Type: Double
Parameter Sets: Rectangle
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

- `System.String`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
