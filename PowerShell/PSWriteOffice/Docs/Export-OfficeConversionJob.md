---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeConversionJob
## SYNOPSIS
Saves a portable conversion job for Studio, the command-line tool and PowerShell.

## SYNTAX
### Directory (Default)
```powershell
Export-OfficeConversionJob [-Path] <string> -InputDirectory <string> -OutputDirectory <string> [-TargetExtension <string>] [-ConversionRouteId <string>] [-Name <string>] [-CheckpointDirectory <string>] [-ConflictPolicy <string>] [-OutputProfile <string>] [-ConversionOptions <OfficeConversionOptions>] [-SourcePasswordEnvironmentVariable <string>] [-PdfPasswordEnvironmentVariable <string>] [-SourceExtension <string[]>] [-NoRecurse] [-MaximumFiles <int>] [-MaximumConcurrency <int>] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-MaximumXmlCharactersInPart <long>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Files
```powershell
Export-OfficeConversionJob [-Path] <string> -InputPath <Object[]> -OutputDirectory <string> [-TargetExtension <string>] [-ConversionRouteId <string>] [-Name <string>] [-CheckpointDirectory <string>] [-ConflictPolicy <string>] [-OutputProfile <string>] [-ConversionOptions <OfficeConversionOptions>] [-SourcePasswordEnvironmentVariable <string>] [-PdfPasswordEnvironmentVariable <string>] [-MaximumFiles <int>] [-MaximumConcurrency <int>] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-MaximumXmlCharactersInPart <long>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saving records intent and never converts documents. Relative paths are based on the saved job's folder. Native renderer callbacks and settings outside the portable contract are rejected instead of omitted.

## EXAMPLES

### EXAMPLE 1
```powershell
Export-OfficeConversionJob -InputDirectory ./source -OutputDirectory ./pdf -CheckpointDirectory ./state -Path ./monthly.officejob.json
            Invoke-OfficeConversionJob -Path ./monthly.officejob.json -Plan
            Invoke-OfficeConversionJob -Path ./monthly.officejob.json -ErrorAction Stop
```


### EXAMPLE 2
```powershell
Get-ChildItem ./reports -Filter *.docx | Export-OfficeConversionJob -OutputDirectory ./pdf -Path ./reports.officejob.json -SourcePasswordEnvironmentVariable DOCUMENT_PASSWORD
```


## PARAMETERS

### -CheckpointDirectory
Optional durable state directory outside source and output folders. Registered adapter routes do not support checkpoints.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConflictPolicy
Existing destination behavior. Durable checkpoint jobs require Fail.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values: Fail, Rename, Replace

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConversionOptions
Portable scalar reconstruction and plain-text settings from New-OfficeConversionOptions. Unsupported native settings are rejected.

```yaml
Type: OfficeConversionOptions
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConversionRouteId
Optional exact executable conversion route identifier.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Replace an existing job file. Source documents and protected execution paths remain protected.

```yaml
Type: SwitchParameter
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputDirectory
Source directory to discover when the job runs.

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
Selected source path strings or FileInfo objects.

```yaml
Type: Object[]
Parameter Sets: Files
Aliases: FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### -MaximumConcurrency
Maximum simultaneously executing conversion items.

```yaml
Type: Int32
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumFiles
Maximum discovered or selected files.

```yaml
Type: Int32
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInputBytes
Maximum input bytes per document.

```yaml
Type: Int64
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumOutputBytes
Maximum output bytes per document.

```yaml
Type: Int64
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumXmlCharactersInPart
Maximum XML characters per part for applicable conversions.

```yaml
Type: Int64
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Optional descriptive job name.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRecurse
Do not discover source subdirectories.

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
Destination directory for converted documents.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputProfile
Shared output intent, subject to each route's supported profiles.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values: Faithful, Lightweight, PrintReady, TextOnly

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Destination JSON file. Existing job files require Force.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: LiteralPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfPasswordEnvironmentVariable
Runtime environment variable supplying a PDF input password. Only the name is saved.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceExtension
Restrict directory discovery to these source extensions.

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

### -SourcePasswordEnvironmentVariable
Runtime environment variable supplying an encrypted Office input password. Only the name is saved.

```yaml
Type: String
Parameter Sets: Directory, Files
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetExtension
Destination format extension, with or without the leading dot.

```yaml
Type: String
Parameter Sets: Directory, Files
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

- `System.Object[]`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
