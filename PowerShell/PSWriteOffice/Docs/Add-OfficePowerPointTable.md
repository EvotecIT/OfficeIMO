---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointTable
## SYNOPSIS
Adds a table to a PowerPoint slide.

## SYNTAX
### InputObject (Default)
```powershell
Add-OfficePowerPointTable [[-Slide] <PowerPointSlide>] [-InputObject] <Object> [-Header <string[]>] [-NoHeader] [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-StyleId <string>] [-PassThru] [<CommonParameters>]
```

### Size
```powershell
Add-OfficePowerPointTable [[-Slide] <PowerPointSlide>] -Rows <int> -Columns <int> [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-StyleId <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Builds a table from data rows or creates a blank grid with a fixed size.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = @([pscustomobject]@{ Item='Alpha'; Qty=2 }, [pscustomobject]@{ Item='Beta'; Qty=4 })
Add-OfficePowerPointTable -Slide $slide -InputObject $rows -X 60 -Y 140 -Width 420 -Height 200
```

Creates a table with headers and two data rows.

## PARAMETERS

### -CollectionSeparator
Text used between items when a cell contains a collection.

```yaml
Type: String
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Columns
Column count for an empty table.

```yaml
Type: Int32
Parameter Sets: Size
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DictionaryEntrySeparator
Text used between entries when a cell contains a dictionary.

```yaml
Type: String
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DictionaryKeyValueSeparator
Text used between a dictionary key and value.

```yaml
Type: String
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Optional header order to apply to the table.

```yaml
Type: String[]
Parameter Sets: InputObject
Aliases: Headers
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Table height in points.

```yaml
Type: Double
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Source objects to convert into table rows.

```yaml
Type: Object
Parameter Sets: InputObject
Aliases: Data
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxCollectionItems
Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.

```yaml
Type: Int32
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxNestingDepth
Maximum nesting depth allowed while normalizing one cell value. Defaults to 64; increase explicitly for trusted deeper values.

```yaml
Type: Int32
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
Skip writing header row.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rows
Row count for an empty table.

```yaml
Type: Int32
Parameter Sets: Size
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Target slide that will receive the table (optional inside DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -StyleId
Optional table style ID (GUID string).

```yaml
Type: String
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -View
Projection to apply before writing the table.

```yaml
Type: OfficeTableView
Parameter Sets: InputObject
Aliases: None
Possible values: Normal, Transpose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Table width in points.

```yaml
Type: Double
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Left offset (in points) from the slide origin.

```yaml
Type: Double
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Top offset (in points) from the slide origin.

```yaml
Type: Double
Parameter Sets: InputObject, Size
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

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `None`

## RELATED LINKS

- None
