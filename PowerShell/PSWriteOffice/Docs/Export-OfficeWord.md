---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeWord
## SYNOPSIS
Exports PowerShell objects to a Word document containing an editable table.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeWord [-Path] <string> -InputObject <Object[]> [-Title <string>] [-Property <string[]>] [-ColumnFormat <hashtable>] [-Orientation <OfficePageOrientation>] [-Style <WordTableStyle>] [-NoHeader] [-NoClobber] [-Show] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Accepts pipeline rows, arrays, dictionaries and tabular data. Uses the shared table normalizer and native Word writer. Replaces an existing DOCX atomically unless NoClobber is supplied, and stays quiet unless PassThru is supplied.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-Service | Select-Object Name, Status, StartType | Export-OfficeWord -Path ./Services.docx -Title 'Services'
```


## PARAMETERS

### -ColumnFormat
Selected column names mapped to .NET display formats, such as @{ Total = 'N0' }. Formatting uses invariant culture.

```yaml
Type: Hashtable
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Objects supplied as an array or through the pipeline.

```yaml
Type: Object[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoClobber
Refuse to replace an existing destination.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
Omit the column-name row and repeating header.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Orientation
Page orientation; omission chooses landscape when first-row words and padding cannot fit the portrait text area.

```yaml
Type: OfficePageOrientation
Parameter Sets: Path
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Destination DOCX file; the parent folder is created only when publishing.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Property
Property names in display order; omission includes all properties.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved document using the system's default application.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
Native table style; Modern includes a repeating first-row header.

```yaml
Type: WordTableStyle
Parameter Sets: Path
Aliases: None
Possible values: TableNormal, TableGrid, PlainTable1, PlainTable2, PlainTable3, PlainTable4, PlainTable5, GridTable1Light, GridTable1LightAccent1, GridTable1LightAccent2, GridTable1LightAccent3, GridTable1LightAccent4, GridTable1LightAccent5, GridTable1LightAccent6, GridTable2, GridTable2Accent1, GridTable2Accent2, GridTable2Accent3, GridTable2Accent4, GridTable2Accent5, GridTable2Accent6, GridTable3, GridTable3Accent1, GridTable3Accent2, GridTable3Accent3, GridTable3Accent4, GridTable3Accent5, GridTable3Accent6, GridTable4, GridTable4Accent1, GridTable4Accent2, GridTable4Accent3, GridTable4Accent4, GridTable4Accent5, GridTable4Accent6, GridTable5Dark, GridTable5DarkAccent1, GridTable5DarkAccent2, GridTable5DarkAccent3, GridTable5DarkAccent4, GridTable5DarkAccent5, GridTable5DarkAccent6, GridTable6Colorful, GridTable6ColorfulAccent1, GridTable6ColorfulAccent2, GridTable6ColorfulAccent3, GridTable6ColorfulAccent4, GridTable6ColorfulAccent5, GridTable6ColorfulAccent6, GridTable7Colorful, GridTable7ColorfulAccent1, GridTable7ColorfulAccent2, GridTable7ColorfulAccent3, GridTable7ColorfulAccent4, GridTable7ColorfulAccent5, GridTable7ColorfulAccent6, ListTable1Light, ListTable1LightAccent1, ListTable1LightAccent2, ListTable1LightAccent3, ListTable1LightAccent4, ListTable1LightAccent5, ListTable1LightAccent6, ListTable2, ListTable2Accent1, ListTable2Accent2, ListTable2Accent3, ListTable2Accent4, ListTable2Accent5, ListTable2Accent6, ListTable3, ListTable3Accent1, ListTable3Accent2, ListTable3Accent3, ListTable3Accent4, ListTable3Accent5, ListTable3Accent6, ListTable4, ListTable4Accent1, ListTable4Accent2, ListTable4Accent3, ListTable4Accent4, ListTable4Accent5, ListTable4Accent6, ListTable5Dark, ListTable5DarkAccent1, ListTable5DarkAccent2, ListTable5DarkAccent3, ListTable5DarkAccent4, ListTable5DarkAccent5, ListTable5DarkAccent6, ListTable6Colorful, ListTable6ColorfulAccent1, ListTable6ColorfulAccent2, ListTable6ColorfulAccent3, ListTable6ColorfulAccent4, ListTable6ColorfulAccent5, ListTable6ColorfulAccent6, ListTable7Colorful, ListTable7ColorfulAccent1, ListTable7ColorfulAccent2, ListTable7ColorfulAccent3, ListTable7ColorfulAccent4, ListTable7ColorfulAccent5, ListTable7ColorfulAccent6, OfficeIMOModern

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional title above the table, styled with the document's Title style.

```yaml
Type: String
Parameter Sets: Path
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
