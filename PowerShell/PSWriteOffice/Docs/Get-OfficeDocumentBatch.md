---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentBatch
## SYNOPSIS
Reads supported files and folders with adjustable concurrency and limits.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentBatch [-Path] <string[]> [-Recurse] [-Extension <string[]>] [-MaxDocuments <Int32>] [-NoDocumentLimit] [-MaxDegreeOfParallelism <Int32>] [-MaxStoreItems <Int32>] [-AllStoreItems] [-IncludePageLocations] [-ContinueOnError] [<CommonParameters>]
```

## DESCRIPTION
Reads supported files and folders with adjustable concurrency and limits.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeDocumentBatch -Path .\Reports -Recurse -MaxDegreeOfParallelism 4 -ContinueOnError
```

PSWriteOffice discovers registered formats and reports individual read failures without requiring .NET option objects.

## PARAMETERS

### -AllStoreItems
Project every matching item from each email store.

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

### -ContinueOnError
Report individual read errors and continue processing other documents.

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

### -Extension
Optional extensions to include. Registered Reader formats are used automatically when omitted.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: Extensions
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludePageLocations
Compute Word and RTF page locations when supported.

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

### -MaxDegreeOfParallelism
Maximum document reads in flight.

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

### -MaxDocuments
Maximum documents accepted in one batch. The default is 500.

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

### -MaxStoreItems
Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.

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

### -NoDocumentLimit
Remove the document-count ceiling.

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

### -Path
File, directory, or wildcard paths to read.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: FullName, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### -Recurse
Search subdirectories when a path names a directory.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentReadResult`

## RELATED LINKS

- None
