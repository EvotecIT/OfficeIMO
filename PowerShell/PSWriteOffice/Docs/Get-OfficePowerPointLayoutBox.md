---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePowerPointLayoutBox
## SYNOPSIS
Computes native content, column or row boxes for manual placement.

## SYNTAX
### Content (Default)
```powershell
Get-OfficePowerPointLayoutBox [-Margin <OfficeLength>] [-Document <PowerPointDocument>] [<CommonParameters>]
```

### Columns
```powershell
Get-OfficePowerPointLayoutBox -Columns <int> [-Margin <OfficeLength>] [-Gutter <OfficeLength>] [-Document <PowerPointDocument>] [<CommonParameters>]
```

### Rows
```powershell
Get-OfficePowerPointLayoutBox -Rows <int> [-Margin <OfficeLength>] [-Gutter <OfficeLength>] [-Document <PowerPointDocument>] [<CommonParameters>]
```

## DESCRIPTION
Computes native content, column or row boxes for manual placement.

## EXAMPLES

### EXAMPLE 1
```powershell
$box = Get-OfficePowerPointLayoutBox -Document $deck -Margin 1.5cm
```


### EXAMPLE 2
```powershell
$columns = Get-OfficePowerPointLayoutBox -Document $deck -Columns 2 -Margin 1.5cm -Gutter 1cm
PowerPointSlide -Document $deck -Title 'Status' {
    PowerPointTextBox 'Healthy' -X $columns[0].Left -Y $columns[0].Top -Width $columns[0].Width -Height 40
}
```


## PARAMETERS

### -Columns
Gets or sets the number of columns.

```yaml
Type: Int32
Parameter Sets: Columns
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: PowerPointDocument
Parameter Sets: Content, Columns, Rows
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Gutter
Gets or sets the space between rows or columns; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Columns, Rows
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margin
Gets or sets the outer margin; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Content, Columns, Rows
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rows
Gets or sets the number of rows.

```yaml
Type: Int32
Parameter Sets: Rows
Aliases: None
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

- `OfficeIMO.PowerPoint.PowerPointDocument`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointLayoutBox`

## RELATED LINKS

- None
