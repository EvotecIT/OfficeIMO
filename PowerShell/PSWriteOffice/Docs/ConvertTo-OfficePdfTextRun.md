---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfTextRun
## SYNOPSIS
Converts reusable Office text run specifications to native PDF text runs.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfTextRun [-Run] <Object> [<CommonParameters>]
```

## DESCRIPTION
Use this adapter when an OfficeIMO PDF callback requires a native PdfTextRun,
such as a rich generated header or footer. Styling remains owned by New-OfficeTextRun.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $label = New-OfficeTextRun -Text 'Service report ' -Bold -Color '#B42318' | ConvertTo-OfficePdfTextRun
$pageStyle = New-OfficeTextRun -Italic | ConvertTo-OfficePdfTextRun
Set-OfficePdfHeader -Compose {
    param($header)
    $null = $header.Text({
        param($text)
        $null = $text.Run($label).CurrentPage($pageStyle)
    })
}
```

The cross-format run specification stays PowerShell-friendly while the callback receives the native PDF run it requires.

## PARAMETERS

### -Run
One or more values accepted by New-OfficeTextRun, including run specifications and hashtables.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `OfficeIMO.Pdf.PdfTextRun`

## RELATED LINKS

- None
