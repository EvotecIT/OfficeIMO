---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfOptimized
## SYNOPSIS
Applies lossless PDF optimization actions and writes a new PDF.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfOptimized [-Path] <string> [-OutputPath] <string> [-Password <string>] [-IgnorePermissionRestrictions] [-NoCompressStreams] [-KeepUnreferencedObjects] [-KeepDuplicateStreams] [-AllowLarger] [-MinimumStreamCompressionBytes <int>] [-PassThruReport] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Applies lossless PDF optimization actions and writes a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficePdfOptimized -Path .\Report.pdf -OutputPath .\Report-Optimized.pdf
```

Writes a smaller PDF when safe lossless optimization actions can reduce the file size.

## PARAMETERS

### -AllowLarger
Write the optimized candidate even when it is not smaller than the source PDF.

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

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed modification restrictions.

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

### -KeepDuplicateStreams
Keep byte-identical stream objects instead of rewriting duplicate references to one object.

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

### -KeepUnreferencedObjects
Keep orphaned indirect PDF objects instead of pruning objects unreachable from the catalog.

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

### -MinimumStreamCompressionBytes
Minimum unfiltered stream size considered for compression.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCompressStreams
Skip Flate compression of unfiltered streams.

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

### -PassThruReport
Return the optimization action report instead of the output file.

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
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Pdf.PdfOptimizationActionResult`

## RELATED LINKS

- None
