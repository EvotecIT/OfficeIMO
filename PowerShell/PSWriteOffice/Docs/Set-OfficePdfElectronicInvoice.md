---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePdfElectronicInvoice
## SYNOPSIS
Configures Factur-X/ZUGFeRD e-invoice groundwork on a generated PDF document.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePdfElectronicInvoice [-Path] <string> [-Profile <PdfComplianceProfile>] [-ConformanceLevel <string>] [-Version <string>] [-Relationship <PdfAssociatedFileRelationship>] [-Description <string>] [<CommonParameters>]
```

## DESCRIPTION
Configures Factur-X/ZUGFeRD e-invoice groundwork on a generated PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\Invoice.pdf {
    Set-OfficePdfElectronicInvoice -Path .\Examples\Documents\factur-x.xml -Profile FacturX
    Add-OfficePdfHeading -Text 'Invoice'
}
```

Embeds the XML as canonical factur-x.xml, emits matching XMP metadata, and configures PDF/A-3 groundwork.

## PARAMETERS

### -ConformanceLevel
Factur-X/ZUGFeRD conformance level written to XMP metadata.

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

### -Description
Optional human-readable attachment description.

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

### -Path
CrossIndustryInvoice XML file path to embed as canonical factur-x.xml.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath, XmlPath, InvoiceXmlPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Profile
E-invoice profile to prepare.

```yaml
Type: PdfComplianceProfile
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, PdfA2B, PdfA2U, PdfA2A, PdfA3B, PdfA3U, PdfA3A, PdfUa1, FacturX, Zugferd, PdfA4, PdfA4E, PdfA4F, PdfUa2, PdfX1A2003, PdfX4

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Relationship
Associated-file relationship for the embedded XML payload.

```yaml
Type: PdfAssociatedFileRelationship
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Unspecified, Source, Data, Alternative, Supplement, C2paManifest

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Version
Factur-X/ZUGFeRD schema version written to XMP metadata.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
