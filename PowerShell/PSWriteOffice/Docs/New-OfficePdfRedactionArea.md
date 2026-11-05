---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfRedactionArea
## SYNOPSIS
Creates a redaction area for Remove-OfficePdfContent and Get-OfficePdfRedactionPlan -Area.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfRedactionArea [-Page] <int> -X <OfficeLength> -Y <OfficeLength> -Width <OfficeLength> -Height <OfficeLength> [-Label <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates a redaction area for Remove-OfficePdfContent and Get-OfficePdfRedactionPlan -Area.

## EXAMPLES

### EXAMPLE 1
```powershell
$area = New-OfficePdfRedactionArea -Page 1 -X 2cm -Y 3cm -Width 8cm -Height 2cm -Label 'Signature'
Remove-OfficePdfContent -Path .\Contracts\Signed.pdf -OutputPath .\Contracts\Signed-Redacted.pdf -Area $area
```


## PARAMETERS

### -Height
Height for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Label
Label for this operation.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
Page for this operation.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
X for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Y for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Default
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

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfRedactionArea`

## RELATED LINKS

- None
