---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentChunk
## SYNOPSIS
Reads supported Office, PDF, Markdown, RTF, HTML, CSV, JSON, XML, YAML, ZIP, EPUB, Visio, and text files into normalized OfficeIMO.Reader chunks.

## SYNTAX
### File (Default)
```powershell
Get-OfficeDocumentChunk [-Path] <string> [-MaxInputBytes <Int64>] [-OpenXmlMaxCharactersInPart <Int64>] [-MaxChars <Int32>] [-MaxTableRows <Int32>] [-ExcludeWordFootnotes] [-ExcludePowerPointNotes] [-NoExcelHeaders] [-ExcelChunkRows <Int32>] [-ExcelSheetName <string>] [-ExcelA1Range <string>] [-NoMarkdownHeadingChunks] [-NoHashes] [-MaxStoreItems <Int32>] [-AllStoreItems] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

### Folder
```powershell
Get-OfficeDocumentChunk -FolderPath <string> [-NoRecurse] [-MaxFiles <Int32>] [-MaxTotalBytes <Int64>] [-Extension <string[]>] [-MaxInputBytes <Int64>] [-OpenXmlMaxCharactersInPart <Int64>] [-MaxChars <Int32>] [-MaxTableRows <Int32>] [-ExcludeWordFootnotes] [-ExcludePowerPointNotes] [-NoExcelHeaders] [-ExcelChunkRows <Int32>] [-ExcelSheetName <string>] [-ExcelA1Range <string>] [-NoMarkdownHeadingChunks] [-NoHashes] [-MaxStoreItems <Int32>] [-AllStoreItems] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

## DESCRIPTION
This is a thin adapter over OfficeDocumentReader. The OfficeIMO.Reader engine owns detection,
extraction, hashing, and chunk shaping.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeRtf -Path .\Report.rtf -Text 'Summary', 'Ready for review'
            Get-OfficeDocumentChunk -Path .\Report.rtf | Select-Object Kind, Text
```

Creates a small RTF file and reads it back through the Reader adapter as normalized chunks.

## PARAMETERS

### -AllStoreItems
Project every matching item from each email store.

```yaml
Type: SwitchParameter
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelA1Range
Optional Excel A1 range to read.

```yaml
Type: String
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelChunkRows
Excel rows per emitted worksheet chunk.

```yaml
Type: Int32
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelSheetName
Optional Excel sheet name to read.

```yaml
Type: String
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludePowerPointNotes
Exclude PowerPoint speaker notes.

```yaml
Type: SwitchParameter
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeWordFootnotes
Exclude Word footnotes.

```yaml
Type: SwitchParameter
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Extension
Allowed folder extensions such as .docx, .xlsx, .pdf, or md.

```yaml
Type: String[]
Parameter Sets: Folder
Aliases: Extensions
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FolderPath
Folder path to read.

```yaml
Type: String
Parameter Sets: Folder
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxChars
Maximum emitted chunk characters.

```yaml
Type: Int32
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxFiles
Maximum number of folder files to read.

```yaml
Type: Int32
Parameter Sets: Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxInputBytes
Maximum input size in bytes.

```yaml
Type: Int64
Parameter Sets: File, Folder
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
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxTableRows
Maximum table rows per emitted table chunk.

```yaml
Type: Int32
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxTotalBytes
Maximum total folder bytes to read.

```yaml
Type: Int64
Parameter Sets: Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoExcelHeaders
Treat the first Excel row as data instead of headers.

```yaml
Type: SwitchParameter
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHashes
Disable source and chunk hash computation.

```yaml
Type: SwitchParameter
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoMarkdownHeadingChunks
Do not split Markdown by headings.

```yaml
Type: SwitchParameter
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRecurse
Do not recurse into child folders when reading a folder.

```yaml
Type: SwitchParameter
Parameter Sets: Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OpenXmlMaxCharactersInPart
OpenXML maximum characters per part.

```yaml
Type: Int64
Parameter Sets: File, Folder
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
File path to read.

```yaml
Type: String
Parameter Sets: File
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Reader
Optional immutable OfficeIMO reader with caller-configured handlers and processors.

```yaml
Type: OfficeDocumentReader
Parameter Sets: File, Folder
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Reader.ReaderChunk`

## RELATED LINKS

- None
