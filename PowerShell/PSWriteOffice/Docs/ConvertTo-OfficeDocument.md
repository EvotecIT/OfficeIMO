---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# ConvertTo-OfficeDocument
## SYNOPSIS
Converts a file or file batch through the canonical OfficeIMO workflow runner.

## SYNTAX
### File (Default)
```powershell
ConvertTo-OfficeDocument [-InputPath] <Object[]> [-OutputPath] <string> [-ConversionRouteId <string>] [-ConflictPolicy <string>] [-OutputProfile <string>] [-ConversionOptions <OfficeConversionOptions>] [-Password <string>] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-SummaryVariable <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Files
```powershell
ConvertTo-OfficeDocument [-InputPath] <Object[]> -OutputDirectory <string> -TargetExtension <string> [-ConversionRouteId <string>] [-ConflictPolicy <string>] [-OutputProfile <string>] [-ConversionOptions <OfficeConversionOptions>] [-Password <string>] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-MaximumFiles <int>] [-MaximumConcurrency <int>] [-CheckpointDirectory <string>] [-RetryFailed] [-ItemResults] [-AllowEmpty] [-SummaryVariable <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Directory
```powershell
ConvertTo-OfficeDocument -InputDirectory <string> -OutputDirectory <string> -TargetExtension <string> [-ConversionRouteId <string>] [-ConflictPolicy <string>] [-OutputProfile <string>] [-ConversionOptions <OfficeConversionOptions>] [-Password <string>] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-MaximumFiles <int>] [-MaximumConcurrency <int>] [-SourceExtension <string[]>] [-NoRecurse] [-CheckpointDirectory <string>] [-RetryFailed] [-ItemResults] [-AllowEmpty] [-SummaryVariable <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Returns verified workflow evidence; unsupported input is diagnosed by the owning route.

## EXAMPLES

### EXAMPLE 1
```powershell
ConvertTo-OfficeDocument -InputPath ./report.pdf -OutputPath ./report.docx
```


### EXAMPLE 2
```powershell
Get-ChildItem ./reports -Filter *.pdf | ConvertTo-OfficeDocument -OutputDirectory ./word -TargetExtension docx -ItemResults
```


### EXAMPLE 3
```powershell
ConvertTo-OfficeDocument -InputDirectory ./source -OutputDirectory ./pdf -TargetExtension pdf -CheckpointDirectory ./state -ConversionOptions (New-OfficeConversionOptions -RequireNoLoss)
```


## PARAMETERS

### -AllowEmpty
Allow a batch that selects no executable inputs to complete without an error.

```yaml
Type: SwitchParameter
Parameter Sets: Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CheckpointDirectory
Separate durable state directory. Completed artifacts are verified and reused.

```yaml
Type: String
Parameter Sets: Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConflictPolicy
Existing destination behavior. Checkpoint jobs require Fail.

```yaml
Type: String
Parameter Sets: File, Files, Directory
Aliases: None
Possible values: Fail, Rename, Replace

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConversionOptions
Renderer and reconstruction settings from New-OfficeConversionOptions.

```yaml
Type: OfficeConversionOptions
Parameter Sets: File, Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConversionRouteId
Optional exact canonical route identifier from Get-OfficeConversionRoute.

```yaml
Type: String
Parameter Sets: File, Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputDirectory
Literal source directory for a batch.

```yaml
Type: String
Parameter Sets: Directory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputPath
Literal source paths or FileInfo objects. Strings, arrays and file pipelines share one bounded selection.

```yaml
Type: Object[]
Parameter Sets: File, Files
Aliases: Path, LiteralPath, FullName
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### -ItemResults
Stream per-file evidence instead of only the bounded batch summary.

```yaml
Type: SwitchParameter
Parameter Sets: Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumConcurrency
Maximum concurrently running batch items.

```yaml
Type: Int32
Parameter Sets: Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumFiles
Maximum selected or discovered files.

```yaml
Type: Int32
Parameter Sets: Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInputBytes
Maximum bytes read from one source.

```yaml
Type: Int64
Parameter Sets: File, Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumOutputBytes
Maximum generated bytes per document.

```yaml
Type: Int64
Parameter Sets: File, Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRecurse
Do not traverse child directories.

```yaml
Type: SwitchParameter
Parameter Sets: Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Destination directory for selected-file or directory batches.

```yaml
Type: String
Parameter Sets: Files, Directory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Single-file destination; its extension selects the target format.

```yaml
Type: String
Parameter Sets: File
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputProfile
Cross-format output intent, subject to the selected route's supported profiles.

```yaml
Type: String
Parameter Sets: File, Files, Directory
Aliases: None
Possible values: Faithful, Lightweight, PrintReady, TextOnly

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Runtime-only password for encrypted Office or PDF source files.

```yaml
Type: String
Parameter Sets: File, Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RetryFailed
Retry recorded failures without replacing verified completed artifacts.

```yaml
Type: SwitchParameter
Parameter Sets: Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceExtension
Limit directory discovery to these source extensions.

```yaml
Type: String[]
Parameter Sets: Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SummaryVariable
Variable receiving the final batch summary or single-file result.

```yaml
Type: String
Parameter Sets: File, Files, Directory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetExtension
Target extension for every batch item, with or without a leading dot.

```yaml
Type: String
Parameter Sets: Files, Directory
Aliases: To
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

- `System.Object[]`

## OUTPUTS

- `OfficeIMO.OfficeConversionReport`
- `OfficeIMO.OfficeBatchSummary`
- `OfficeIMO.OfficeBatchItemReport`

## RELATED LINKS

- None
