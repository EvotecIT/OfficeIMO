---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfCompliance
## SYNOPSIS
Sets generated PDF compliance profile and readiness groundwork.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfCompliance -Profile <PdfComplianceProfile> [-Groundwork] [-Language <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfCompliance -Document <PdfDocument> -Profile <PdfComplianceProfile> [-Groundwork] [-Language <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets generated PDF compliance profile and readiness groundwork.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $path = '.\Examples\Documents\PdfCompliance.pdf'
New-OfficePdf -Path $path {
    Set-OfficePdfCompliance -Profile PdfA3B -Groundwork -Language 'en-US'
    Add-OfficePdfHeading -Text 'Compliance-ready report'
}
Get-OfficePdfCompliance -Path $path -Profile PdfA3B
```

Applies OfficeIMO.Pdf compliance groundwork during composition, saves the PDF, then inspects the saved file.

## PARAMETERS

### -Document
Compatibility parameter. Compliance options must be declared inside New-OfficePdf.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Groundwork
Configure common PDF/A or PDF/UA groundwork for the selected profile.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Language
Catalog language used by compliance groundwork.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Profile
Requested generated PDF compliance profile.

```yaml
Type: PdfComplianceProfile
Parameter Sets: Context, Document
Aliases: None
Possible values: None, PdfA2B, PdfA2U, PdfA2A, PdfA3B, PdfA3U, PdfA3A, PdfUa1, FacturX, Zugferd, PdfA4, PdfA4E, PdfA4F, PdfUa2, PdfX1A2003, PdfX4

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
