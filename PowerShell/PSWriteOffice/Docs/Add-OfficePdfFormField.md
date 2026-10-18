---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfFormField
## SYNOPSIS
Adds a simple AcroForm field to a generated PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfFormField [-Name] <string> [-Type <OfficePdfFormFieldType>] [-Value <string>] [-Values <string[]>] [-Options <string[]>] [-Checked] [-Width <double>] [-Height <double>] [-Align <PdfAlign>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfFormField [-Name] <string> -Document <PdfDocument> [-Type <OfficePdfFormFieldType>] [-Value <string>] [-Values <string[]>] [-Options <string[]>] [-Checked] [-Width <double>] [-Height <double>] [-Align <PdfAlign>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a simple AcroForm field to a generated PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfForm.pdf {
  Add-OfficePdfHeading -Text 'Access request'
  Add-OfficePdfParagraph -Text 'Requester'
  Add-OfficePdfFormField -Name 'Requester' -Type Text -Width 240
  Add-OfficePdfParagraph -Text 'Priority'
  Add-OfficePdfFormField -Name 'Priority' -Type Choice -Options 'Low','Normal','High' -Value 'Normal'
  Add-OfficePdfFormField -Name 'Approved' -Type CheckBox
}
```

Adds text, choice, and checkbox form fields to a generated PDF.

## PARAMETERS

### -Align
Field alignment.

```yaml
Type: PdfAlign
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right, Justify

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Checked
Initial check-box state.

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

### -Document
PDF document to update outside the DSL context.

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

### -Height
Rendered field height in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Form field name.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Choice or radio options.

```yaml
Type: String[]
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

### -Type
Field type to add.

```yaml
Type: OfficePdfFormFieldType
Parameter Sets: Context, Document
Aliases: None
Possible values: Text, CheckBox, Choice, MultiSelectChoice, RadioButton

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Initial text, selected choice, or selected radio value.

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

### -Values
Initial selected values for multi-select choice fields.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Rendered field width in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
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

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
