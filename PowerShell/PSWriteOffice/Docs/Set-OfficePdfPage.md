---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfPage
## SYNOPSIS
Sets page-level PDF properties and writes a new PDF.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePdfPage -Path <string> -OutputPath <string> [-PageRange <string>] [-Rotation <int>] [-BoxName <PdfPageBoundaryBox>] [-Left <Double>] [-Bottom <Double>] [-Right <Double>] [-Top <Double>] [-PageSize <string>] [-Width <Double>] [-Height <Double>] [-Landscape] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <Double>] [-Password <string>] [-IgnorePermissionRestrictions] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Sets page-level PDF properties and writes a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $proof = @(
    Set-OfficePdfPage -Path .\Examples\Documents\Scanned.pdf -PageRange '2,4' -Rotation 90 -OutputPath .\Examples\Documents\Scanned-Rotated.pdf
    Get-OfficePdfInfo -Path .\Examples\Documents\Scanned-Rotated.pdf | Select-Object PageCount
)
$proof
```

Rotates selected pages and writes a new PDF.

## PARAMETERS

### -Bottom
Bottom coordinate for the page boundary box.

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

### -BoxName
Page boundary box to set. Supported values are MediaBox, CropBox, BleedBox, TrimBox, and ArtBox.

```yaml
Type: PdfPageBoundaryBox
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: MediaBox, CropBox, BleedBox, TrimBox, ArtBox

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Custom page height in points when -PageSize Custom is used.

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

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed page-modification restrictions.

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

### -Landscape
Use the landscape orientation of the selected page size.

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

### -Left
Left coordinate for the page boundary box.

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
Page ranges such as 1-3,5. Omit to affect all pages.

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

### -PageSize
Resize selected pages to a known OfficeIMO page size such as A4, Letter, or Custom.

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
Password used to authenticate an encrypted PDF.

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

### -ResizeMargin
Margin, in points, reserved around resized page content.

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

### -ResizeMode
How source page content is fitted into the resized output page.

```yaml
Type: PdfPageResizeMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Fit, Fill, Stretch

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Right
Right coordinate for the page boundary box.

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

### -Rotation
Rotation in degrees. Supported values are 0, 90, 180, and 270.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 0, 90, 180, 270

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Top
Top coordinate for the page boundary box.

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

### -Width
Custom page width in points when -PageSize Custom is used.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
