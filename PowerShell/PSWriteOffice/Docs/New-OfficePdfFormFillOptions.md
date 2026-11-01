---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfFormFillOptions
## SYNOPSIS
Creates native XFDF field/value limits and appearance behavior for Import-OfficePdfXfdf.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfFormFillOptions [-MaxFields <Int32>] [-MaxValueCharacters <Int32>] [-KeepNeedAppearances] [<CommonParameters>]
```

## DESCRIPTION
Creates native XFDF field/value limits and appearance behavior for Import-OfficePdfXfdf.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePdfFormFillOptions -MaxFields 100 -MaxValueCharacters 10000
Import-OfficePdfXfdf -Path ./Form.pdf -XfdfPath ./Values.xfdf -OutputPath ./Filled.pdf -Options $options
```


## PARAMETERS

### -KeepNeedAppearances
Keep the native NeedAppearances flag set for legacy viewers.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxFields
Maximum fields admitted from one XFDF document.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxValueCharacters
Maximum aggregate XFDF value characters.

```yaml
Type: Int32
Parameter Sets: Default
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

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfFormFillerOptions`

## RELATED LINKS

- None
