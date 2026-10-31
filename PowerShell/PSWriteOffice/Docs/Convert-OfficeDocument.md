---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Convert-OfficeDocument
## SYNOPSIS
Converts files through OfficeIMO's native routes and portable batch engine.

## SYNTAX
### Path (Default)
```powershell
Convert-OfficeDocument [-Path] <string[]> -To <string> [-Plan] [-ReportPath <string>] [-Recurse] [-Extension <string[]>] [-MaximumFiles <int>] [-ThrottleLimit <int>] [-Resume] [-RetryFailed] [-CheckpointPath <string>] [-AllowEmpty] [-SaveJob <string>] [-JobName <string>] [-PasswordVariable <string>] [-PdfPasswordVariable <string>] [-OutputDirectory <string>] [-Options <Object[]>] [-Password <securestring>] [-Route <string>] [-OutputProfile <OfficeConversionProfile>] [-FailOnLoss] [-AllowLegacyImportLoss] [-CompressPdf] [-NoClobber] [-SkipUpToDate] [-AutoRename] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-MaximumXmlCharactersInPart <long>] [-SummaryVariable <string>] [-From <string>] [-AsString] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Job
```powershell
Convert-OfficeDocument -JobPath <string> [-Plan] [-ReportPath <string>] [-RetryFailed] [-AllowEmpty] [-SummaryVariable <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Convert-OfficeDocument -LiteralPath <string[]> -To <string> [-Plan] [-ReportPath <string>] [-Recurse] [-Extension <string[]>] [-MaximumFiles <int>] [-ThrottleLimit <int>] [-Resume] [-RetryFailed] [-CheckpointPath <string>] [-AllowEmpty] [-SaveJob <string>] [-JobName <string>] [-PasswordVariable <string>] [-PdfPasswordVariable <string>] [-OutputDirectory <string>] [-Options <Object[]>] [-Password <securestring>] [-Route <string>] [-OutputProfile <OfficeConversionProfile>] [-FailOnLoss] [-AllowLegacyImportLoss] [-CompressPdf] [-NoClobber] [-SkipUpToDate] [-AutoRename] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-MaximumXmlCharactersInPart <long>] [-SummaryVariable <string>] [-From <string>] [-AsString] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathFile
```powershell
Convert-OfficeDocument [-Path] <string[]> [-OutputPath] <string> [-To <string>] [-Options <Object[]>] [-Password <securestring>] [-Route <string>] [-OutputProfile <OfficeConversionProfile>] [-FailOnLoss] [-AllowLegacyImportLoss] [-CompressPdf] [-NoClobber] [-SkipUpToDate] [-AutoRename] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-MaximumXmlCharactersInPart <long>] [-SummaryVariable <string>] [-From <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathFile
```powershell
Convert-OfficeDocument [-OutputPath] <string> -LiteralPath <string[]> [-To <string>] [-Options <Object[]>] [-Password <securestring>] [-Route <string>] [-OutputProfile <OfficeConversionProfile>] [-FailOnLoss] [-AllowLegacyImportLoss] [-CompressPdf] [-NoClobber] [-SkipUpToDate] [-AutoRename] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-MaximumXmlCharactersInPart <long>] [-SummaryVariable <string>] [-From <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
Convert-OfficeDocument [-Text] <string> -To <string> -From <string> [-OutputPath <string>] [-Options <Object[]>] [-Password <securestring>] [-Route <string>] [-OutputProfile <OfficeConversionProfile>] [-FailOnLoss] [-AllowLegacyImportLoss] [-CompressPdf] [-NoClobber] [-AutoRename] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-MaximumXmlCharactersInPart <long>] [-SummaryVariable <string>] [-AsString] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Writes each completed file's conversion evidence as it finishes. Paths resolve through the FileSystem provider before background execution; LiteralPath preserves wildcard characters. Saved jobs use the same contract as Studio and the CLI.

## EXAMPLES

### EXAMPLE 1
```powershell
Convert-OfficeDocument -Path ./report.docx -OutputPath ./report.pdf
```


### EXAMPLE 2
```powershell
Get-ChildItem ./reports -Filter *.docx | Convert-OfficeDocument -To Pdf -OutputDirectory ./pdf -SummaryVariable summary
            $summary
```


### EXAMPLE 3
```powershell
Convert-OfficeDocument -Path ./reports -Recurse -To Pdf -OutputDirectory ./pdf -SaveJob ./reports.officejob.json
            Convert-OfficeDocument -JobPath ./reports.officejob.json -Plan
```


## PARAMETERS

### -AllowEmpty
Accept an intentionally empty selection without an error.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Job, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowLegacyImportLoss
Explicitly accept the native legacy Word importer's reported reductions.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsString
Return converted text to the pipeline instead of publishing a main output file.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoRename
Choose a unique destination name when the requested output exists.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CheckpointPath
Checkpoint manifest file. With Resume, defaults to .officeimo-batch.json in the output folder.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CompressPdf
Apply verified lossless compression to PDF output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Extension
Select these source extensions from directories.

```yaml
Type: String[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FailOnLoss
Reject a conversion that reports fidelity loss before publishing its output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -From
Source format family or exact codec extension. Overrides filenames while retaining their physical source identity.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -JobName
Display name stored in the saved definition.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -JobPath
Portable job definition to run or preview; relative paths resolve against this file.

```yaml
Type: String
Parameter Sets: Job
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Input files or folders without wildcard expansion; accepts FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath, LiteralPathFile
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MaximumFiles
Maximum discovered files, including skipped inputs; zero selects the largest supported bound.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInputBytes
Maximum source bytes per file.

```yaml
Type: Int64
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumOutputBytes
Maximum serialized output bytes per file.

```yaml
Type: Int64
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumXmlCharactersInPart
Maximum XML characters in one package part.

```yaml
Type: Int64
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
Leave an existing destination unchanged and report its input as skipped.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Format-owned settings, typed conversion options, or property hashtables.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Destination folder for a batch.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Single-file destination. Its extension selects the target when To is omitted.

```yaml
Type: String
Parameter Sets: PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputProfile
Output intent, subject to the selected route's supported profiles.

```yaml
Type: OfficeConversionProfile
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values: Faithful, Lightweight, PrintReady, TextOnly

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Runtime source password; saved definitions contain references rather than password values.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PasswordVariable
Admitted environment-variable reference for source passwords in a saved job.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Input files or folders, with FileSystem provider wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path, PathFile
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -PdfPasswordVariable
Separate environment-variable reference for PDF source passwords in a mixed-format job.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Plan
Preview selected inputs and proposed destinations without output or checkpoint writes.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Job, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Recurse
Traverse child directories when selecting a folder.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReportPath
New durable run report outside the source, output, checkpoint and job trees.

```yaml
Type: String
Parameter Sets: Path, Job, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Resume
Continue a checkpointed batch, verifying completed output hashes before reuse.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RetryFailed
Retry failed checkpoint items while retaining verified completed outputs.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Job, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Route
Exact executable route identifier from Get-OfficeDocumentCapability -Conversion.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SaveJob
Save the portable batch definition instead of executing it; returns its FileInfo.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipUpToDate
Skip a file whose existing output has a strictly newer UTC modification time.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SummaryVariable
Final batch counts or single-file evidence assigned before any terminal batch error.

```yaml
Type: String
Parameter Sets: Path, Job, LiteralPath, PathFile, LiteralPathFile, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Literal source content, interpreted according to From rather than as a filename.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ThrottleLimit
Maximum simultaneous conversions.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -To
Target format family, such as Pdf or Word, or an exact physical extension, such as .doc or .docx.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, PathFile, LiteralPathFile, Text
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

- `System.String[]`
- `System.String`

## OUTPUTS

- `OfficeIMO.OfficeConversionReport`
- `OfficeIMO.OfficeBatchItemReport`
- `OfficeIMO.OfficeBatchPlan`
- `System.IO.FileInfo`
- `System.String`

## RELATED LINKS

- None
