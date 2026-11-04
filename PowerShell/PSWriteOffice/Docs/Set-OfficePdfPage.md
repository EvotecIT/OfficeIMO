---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePdfPage
## SYNOPSIS
Rotates pages, sets boundary boxes or resizes pages through the native PDF engine.

## SYNTAX
### Document (Default)
```powershell
Set-OfficePdfPage [-Page <OfficePageRange>] [-Rotation <int>] [-BoxName <PdfPageBoundaryBox>] [-BoxLeft <OfficeLength>] [-BoxBottom <OfficeLength>] [-BoxRight <OfficeLength>] [-BoxTop <OfficeLength>] [-PageSize <OfficePageSize>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <OfficeLength>] [-Document <PdfDocument>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficePdfPage [-Path] <string[]> -OutputPath <string> [-Page <OfficePageRange>] [-Rotation <int>] [-BoxName <PdfPageBoundaryBox>] [-BoxLeft <OfficeLength>] [-BoxBottom <OfficeLength>] [-BoxRight <OfficeLength>] [-BoxTop <OfficeLength>] [-PageSize <OfficePageSize>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <OfficeLength>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Set-OfficePdfPage [-Path] <string[]> -OutputDirectory <string> [-Page <OfficePageRange>] [-Rotation <int>] [-BoxName <PdfPageBoundaryBox>] [-BoxLeft <OfficeLength>] [-BoxBottom <OfficeLength>] [-BoxRight <OfficeLength>] [-BoxTop <OfficeLength>] [-PageSize <OfficePageSize>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <OfficeLength>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Set-OfficePdfPage -LiteralPath <string[]> -OutputPath <string> [-Page <OfficePageRange>] [-Rotation <int>] [-BoxName <PdfPageBoundaryBox>] [-BoxLeft <OfficeLength>] [-BoxBottom <OfficeLength>] [-BoxRight <OfficeLength>] [-BoxTop <OfficeLength>] [-PageSize <OfficePageSize>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <OfficeLength>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Set-OfficePdfPage -LiteralPath <string[]> -OutputDirectory <string> [-Page <OfficePageRange>] [-Rotation <int>] [-BoxName <PdfPageBoundaryBox>] [-BoxLeft <OfficeLength>] [-BoxBottom <OfficeLength>] [-BoxRight <OfficeLength>] [-BoxTop <OfficeLength>] [-PageSize <OfficePageSize>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <OfficeLength>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Rotates pages, sets boundary boxes or resizes pages through the native PDF engine.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePdfPage -BoxBottom 'Value'
```


### EXAMPLE 2
```powershell
Set-OfficePdfPage -LiteralPath @('C:\Path') -OutputPath 'C:\Path'
```


### EXAMPLE 3
```powershell
Set-OfficePdfPage -LiteralPath @('C:\Path') -OutputDirectory 'Value'
```


## PARAMETERS

### -BoxBottom
Bottom boundary coordinate.

```yaml
Type: OfficeLength
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BoxLeft
Left boundary coordinate.

```yaml
Type: OfficeLength
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BoxName
Page boundary box to set.

```yaml
Type: PdfPageBoundaryBox
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values: MediaBox, CropBox, BleedBox, TrimBox, ArtBox

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BoxRight
Right boundary coordinate.

```yaml
Type: OfficeLength
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BoxTop
Top boundary coordinate.

```yaml
Type: OfficeLength
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Explicit native input; omission uses the enclosing New-/Open- content block.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Height
Custom page height.

```yaml
Type: OfficeLength
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After authentication, explicitly ignore owner-imposed usage restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal input files, including FileInfo.FullName.

```yaml
Type: String[]
Parameter Sets: LiteralPath, LiteralPathDirectory
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
Native source parsing, authentication and resource settings.

```yaml
Type: PdfLoadOptions
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Orientation
Orientation of the target page.

```yaml
Type: OfficePageOrientation
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Destination folder retaining input file names.

```yaml
Type: String
Parameter Sets: PathDirectory, LiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Single-file destination.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
One-based pages; omission selects every page.

```yaml
Type: OfficePageRange
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageSize
Paper preset or physical dimensions for resizing.

```yaml
Type: OfficePageSize
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated native document or each saved FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Source password.

```yaml
Type: SecureString
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Input files with FileSystem wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path, PathDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ResizeMargin
Margin reserved around fitted content.

```yaml
Type: OfficeLength
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ResizeMode
How original page content fits the resized page.

```yaml
Type: PdfPageResizeMode
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values: Fit, Fill, Stretch

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rotation
Absolute rotation in degrees.

```yaml
Type: Int32
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values: 0, 90, 180, 270

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Custom page width.

```yaml
Type: OfficeLength
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
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

- `OfficeIMO.Pdf.PdfDocument`
- `System.String[]`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
