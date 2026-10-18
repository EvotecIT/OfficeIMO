---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfRedactionPlan
## SYNOPSIS
Previews explicit or text-matched PDF redaction areas before content is removed.

## SYNTAX
### Rectangle (Default)
```powershell
Get-OfficePdfRedactionPlan [-Path] <string> -PageNumber <int> -X <double> -Y <double> -Width <double> -Height <double> [-Label <string>] [-Password <string>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

### Area
```powershell
Get-OfficePdfRedactionPlan [-Path] <string> -Area <PdfRedactionArea[]> [-Password <string>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

### Text
```powershell
Get-OfficePdfRedactionPlan [-Path] <string> -Text <string[]> [-MatchCase] [-Password <string>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

## DESCRIPTION
This command reports redaction impact only. It does not remove or rewrite PDF content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePdfRedactionPlan -Path .\Report.pdf -PageNumber 1 -X 72 -Y 650 -Width 240 -Height 32
```

Returns line-level text blocks and annotations that intersect the rectangle.

### EXAMPLE 2
```powershell
PS> Get-OfficePdfRedactionPlan -Path .\Report.pdf -Text 'Account number'
```

Derives reviewable redaction areas from logical text blocks containing the supplied text.

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
After successful password authentication, explicitly ignore owner-imposed extraction restrictions.

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
PDF file path.

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

- `OfficeIMO.Pdf.PdfRedactionPlan`

## RELATED LINKS

- None
