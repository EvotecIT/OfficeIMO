---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficePdf
## SYNOPSIS
Saves an OfficeIMO.Pdf document.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficePdf [-Document] <PdfDocument> [-Path] <string> [-Open] [-PassThru] [-Password <string>] [-OwnerPassword <string>] [-Permission <Int32>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use this command when a PDF is built in memory and saved later, or when a pipeline should continue with the saved file.
The document is saved through the normal OfficeIMO.Pdf save path.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $pdf = New-OfficePdf { PdfHeading 'Queued report'; PdfParagraph 'Generated in memory.' }
$pdf | Save-OfficePdf -Path .\QueuedReport.pdf
```

Creates a PDF document object first, then saves it to disk.

## PARAMETERS

### -Document
PDF document to save.

```yaml
Type: PdfDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Open
Open the PDF after saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OwnerPassword
Optional owner password for the generated encrypted PDF.

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

### -PassThru
Emit the document for further processing.

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

### -Password
Password required to open the generated PDF.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: UserPassword
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Destination PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Permission
Raw PDF Standard security permission bit mask. Defaults to allowing all standard operations.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: Permissions
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
