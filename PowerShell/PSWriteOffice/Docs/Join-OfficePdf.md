---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Join-OfficePdf
## SYNOPSIS
Joins multiple PDF files into a single PDF.

## SYNTAX
### __AllParameterSets
```powershell
Join-OfficePdf [-Path] <string[]> [-OutputPath] <string> [-Password <string[]>] [-IgnorePermissionRestrictions] [-PassThru] [-PassThruReport] [-FlattenVisualAnnotations] [-PageSize <string>] [-Width <Double>] [-Height <Double>] [-Landscape] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <Double>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Joins multiple PDF files into a single PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $cover = '.\Examples\Documents\Cover.pdf'
$report = '.\Examples\Documents\Report.pdf'
Join-OfficePdf -Path $cover, $report -OutputPath .\Examples\Documents\Combined.pdf -PassThru
Get-OfficePdfInfo -Path .\Examples\Documents\Combined.pdf | Select-Object PageCount
```

Writes a single PDF containing the input documents in the requested order, then checks the result.

### EXAMPLE 2
```powershell
PS> $result = Join-OfficePdf -Path .\Restricted.pdf, .\Appendix.pdf `
    -Password 'source-password', $null -IgnorePermissionRestrictions `
    -OutputPath .\Combined.pdf -PassThruReport
$result.Sources | Select-Object SourceIndex, PasswordAuthenticationRole, PermissionRestrictionsIgnored
$result.Decisions | Select-Object Structure, Mode, Action
```

A valid password remains mandatory. The explicit switch ignores usage flags after authentication and the report records every source decision.

## PARAMETERS

### -FlattenVisualAnnotations
Flatten visual annotation appearances before merging.

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

### -Height
Custom output page height in points when -PageSize Custom is used.

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
After successful password authentication, explicitly ignore owner-imposed usage restrictions such as copying or assembly.
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

### -Landscape
Use the landscape orientation of the selected output page size.

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

### -OutputPath
Output PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageSize
Resize each merged page to a known OfficeIMO page size such as A4, Letter, or Custom.

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
Emit the saved file.

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

### -PassThruReport
Emit the per-source merge inventory, permission decisions, and output security readback.

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
Passwords used to authenticate encrypted sources. Supply one value for every source, or one value to reuse for all sources.

```yaml
Type: String[]
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
Input PDF paths in output order.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
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

### -Width
Custom output page width in points when -PageSize Custom is used.

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
- `OfficeIMO.Pdf.PdfMergeReport`

## RELATED LINKS

- None
