---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordCoverPage
## SYNOPSIS
Adds a built-in cover page template to a Word document.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordCoverPage [-Template] <WordCoverPageTemplate> [-Document <WordDocument>] [-PublishDate <string>] [-Abstract <string>] [-CompanyAddress <string>] [-CompanyEmail <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Uses OfficeIMO.Word cover page templates and optional cover-page metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx { Add-OfficeWordCoverPage -Template Element -Abstract 'Executive summary' }
```

Creates a document with a template-driven cover page.

## PARAMETERS

### -Abstract
Abstract/summary stored in the cover page properties custom XML part.

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

### -CompanyAddress
Company address stored in the cover page properties custom XML part.

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

### -CompanyEmail
Company email stored in the cover page properties custom XML part.

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

### -Document
Document to update. Defaults to the current Word DSL or tracked document.

```yaml
Type: WordDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the created cover page.

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

### -PublishDate
Publish date stored in the cover page properties custom XML part.

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

### -Template
Cover page template to insert.

```yaml
Type: WordCoverPageTemplate
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Austin, Banded, Facet, Grid, IonDark, IonLight, Element, Wisp, ViewMaster, SliceLight, SliceDark, SideLine, Semaphore, Retrospect

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordCoverPage`

## RELATED LINKS

- None
