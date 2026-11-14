---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeWordBarcode
## SYNOPSIS
Adds a QR code or barcode to a Word paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordBarcode [-Text] <string> [-Paragraph <WordParagraph>] [-PassThru] [-Type <SymbolFormat>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-QuietZone <Int32>] [-Color <OfficeColor>] [-BackgroundColor <OfficeColor>] [-ErrorCorrection <QrErrorCorrectionLevel>] [-NoHumanReadableText] [-AlternativeText <string>] [<CommonParameters>]
```

## DESCRIPTION
The code is an SVG picture with a PNG fallback, so Word draws it sharply and older readers still show it; OfficeIMO's PDF conversion places it at print resolution. Two-dimensional codes default to 1.25 inches wide; linear barcodes to 0.33 mm per bar module, with their text underneath. The alternative text defaults to the symbology and the encoded text.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeWord -Path .\Flyer.docx {
    WordParagraph 'Scan for the event agenda:'
    WordBarcode 'https://evotec.xyz/events'
}
```


### EXAMPLE 2
```powershell
WordParagraph { WordText 'Order: '; WordBarcode 'ORD-2026-0042' -Type Code128 -Width 6cm }
```


## PARAMETERS

### -AlternativeText
Accessible description; defaults to the symbology and the encoded text.

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

### -BackgroundColor
Background and quiet zone color; a literal color.

```yaml
Type: OfficeColor
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Color of the bars or dark modules; a literal color such as Black or #1F3864.

```yaml
Type: OfficeColor
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ErrorCorrection
QR error correction level: L, M, Q or H.

```yaml
Type: QrErrorCorrectionLevel
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: L, M, Q, H

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Height of the whole symbol including human-readable text; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHumanReadableText
Leave out the human-readable text that linear barcodes print under the bars.

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

### -Paragraph
Target paragraph; the current DSL paragraph, or a new one, is used when omitted.

```yaml
Type: WordParagraph
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Return the native editable WordImage.

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

### -QuietZone
Light margin around the symbol, in modules.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text, URL or other data to encode.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Symbology; QR code by default. Any CodeGlyphX format, such as Code128, Ean, UpcA, DataMatrix or Pdf417.

```yaml
Type: SymbolFormat
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: QrCode, MicroQrCode, Aztec, Code128, Gs1Code128, Code39, Code93, Ean, UpcA, UpcE, Itf14, Itf, Industrial2Of5, Matrix2Of5, Iata2Of5, PatchCode, Codabar, Msi, Code11, Plessey, Telepen, Pharmacode, PharmacodeTwoTrack, Code32, Postnet, Planet, RoyalMail4State, AustraliaPost, JapanPost, Gs1DataBarTruncated, Gs1DataBarOmnidirectional, Gs1DataBarStacked, Gs1DataBarExpanded, Gs1DataBarExpandedStacked, UspsIntelligentMail, KixCode, DataMatrix, Pdf417, MicroPdf417, RmQrCode, Gs1DataBarLimited, Gs1DataBarStackedOmnidirectional, MaxiCode, DotCode, HanXin, Gs1Composite, Ean2, Ean5, TelepenNumeric, Coop2Of5, Datalogic2Of5, Bc412, FacingIdentificationMark, Daft, Flattermarken, ChannelCode, Code16K, Code49, CodablockF, PosiCode, Mailmark4State, D3Aqr, AztecRune, CodeOne, Ultracode

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width of the whole symbol including its quiet zone; numbers mean points.

```yaml
Type: OfficeLength
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

- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordImage`

## RELATED LINKS

- None
