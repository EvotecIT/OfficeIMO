---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfStamp
## SYNOPSIS
Adds a text or image stamp to an existing PDF.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePdfStamp -Path <string> -OutputPath <string> -Text <string> [-Password <string>] [-IgnorePermissionRestrictions] [-PageRange <string>] [-X <Double>] [-Y <Double>] [-FontSize <double>] [-Color <string>] [-Rotation <double>] [-Watermark] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Image
```powershell
Add-OfficePdfStamp -Path <string> -OutputPath <string> -Image <string> [-Password <string>] [-IgnorePermissionRestrictions] [-PageRange <string>] [-X <Double>] [-Y <Double>] [-Width <Double>] [-Height <Double>] [-Rotation <double>] [-Watermark] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Stamps are existing-PDF operations. Use text stamps for review labels and image stamps for logos or approval marks.
Use -Watermark when the stamp should be placed behind existing page content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $proof = @(
    Add-OfficePdfStamp -Path .\Examples\Documents\Report.pdf -OutputPath .\Examples\Documents\Stamped.pdf -Text 'REVIEWED' -Color '#0F766E' -FontSize 24 -Rotation 12 -PageRange '1-2'
    Get-OfficePdfPreflight -Path .\Examples\Documents\Stamped.pdf
)
$proof
```

Adds a text stamp to the first two pages and preflights the result.

### EXAMPLE 2
```powershell
PS> $logo = '.\Tests\Assets\CellImage.png'
Add-OfficePdfStamp -Path .\Examples\Documents\Report.pdf -OutputPath .\Examples\Documents\Watermarked.pdf -Image $logo -Width 160 -Watermark
```

Adds a logo behind existing content as a watermark.

## PARAMETERS

### -Color
Text color. Named colors and hexadecimal values are accepted.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontSize
Font size for text stamps.

```yaml
Type: Double
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Rendered image height in PDF points.

```yaml
Type: Double
Parameter Sets: Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed usage restrictions.
This does not discover, bypass, or crack a missing password.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Image
Image path to stamp.

```yaml
Type: String
Parameter Sets: Image
Aliases: ImagePath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Output PDF path.

```yaml
Type: String
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageRange
Stamp selected pages, for example 1-3,5. Omit to stamp every page.

```yaml
Type: String
Parameter Sets: Text, Image
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
Parameter Sets: Text, Image
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
Parameter Sets: Text, Image
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
Parameter Sets: Text, Image
Aliases: FilePath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rotation
Rotation in degrees.

```yaml
Type: Double
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text to stamp.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Watermark
Place the stamp behind existing content as a watermark.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Rendered image width in PDF points.

```yaml
Type: Double
Parameter Sets: Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
X coordinate in PDF points.

```yaml
Type: Double
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Y coordinate in PDF points.

```yaml
Type: Double
Parameter Sets: Text, Image
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

- `System.IO.FileInfo`

## RELATED LINKS

- None
