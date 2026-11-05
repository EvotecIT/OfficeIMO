---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfStamp
## SYNOPSIS
Stamps text, images, native canvas content or a PDF page onto existing PDFs.

## SYNTAX
### TextDocument (Default)
```powershell
Add-OfficePdfStamp -Text <string> [-Size <OfficeLength>] [-Color <OfficeColor>] [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-PassThru] [-Document <PdfDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
Add-OfficePdfStamp [-Path] <string[]> -OutputPath <string> -Text <string> [-Size <OfficeLength>] [-Color <OfficeColor>] [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### TextDirectory
```powershell
Add-OfficePdfStamp [-Path] <string[]> -Text <string> -OutputDirectory <string> [-Size <OfficeLength>] [-Color <OfficeColor>] [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Image
```powershell
Add-OfficePdfStamp [-Path] <string[]> -OutputPath <string> -ImagePath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ImageDirectory
```powershell
Add-OfficePdfStamp [-Path] <string[]> -OutputDirectory <string> -ImagePath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePdfStamp [-Path] <string[]> -OutputPath <string> -Content <scriptblock> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-ConfigureRendering <scriptblock>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ContentDirectory
```powershell
Add-OfficePdfStamp [-Path] <string[]> -OutputDirectory <string> -Content <scriptblock> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-ConfigureRendering <scriptblock>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Page
```powershell
Add-OfficePdfStamp [-Path] <string[]> -OutputPath <string> -OverlayPath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-OverlayPage <int>] [-Fit <PdfPageOverlayFit>] [-HorizontalAlignment <OfficeAlignment>] [-VerticalAlignment <OfficeTextVerticalAlignment>] [-OverlayPassword <securestring>] [-IgnoreOverlayPermissionRestrictions] [-OverlayLoadOptions <PdfLoadOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PageDirectory
```powershell
Add-OfficePdfStamp [-Path] <string[]> -OutputDirectory <string> -OverlayPath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-OverlayPage <int>] [-Fit <PdfPageOverlayFit>] [-HorizontalAlignment <OfficeAlignment>] [-VerticalAlignment <OfficeTextVerticalAlignment>] [-OverlayPassword <securestring>] [-IgnoreOverlayPermissionRestrictions] [-OverlayLoadOptions <PdfLoadOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### TextLiteralPath
```powershell
Add-OfficePdfStamp -OutputPath <string> -Text <string> -LiteralPath <string[]> [-Size <OfficeLength>] [-Color <OfficeColor>] [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ImageLiteralPath
```powershell
Add-OfficePdfStamp -OutputPath <string> -LiteralPath <string[]> -ImagePath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ContentLiteralPath
```powershell
Add-OfficePdfStamp -OutputPath <string> -LiteralPath <string[]> -Content <scriptblock> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-ConfigureRendering <scriptblock>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PageLiteralPath
```powershell
Add-OfficePdfStamp -OutputPath <string> -LiteralPath <string[]> -OverlayPath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-OverlayPage <int>] [-Fit <PdfPageOverlayFit>] [-HorizontalAlignment <OfficeAlignment>] [-VerticalAlignment <OfficeTextVerticalAlignment>] [-OverlayPassword <securestring>] [-IgnoreOverlayPermissionRestrictions] [-OverlayLoadOptions <PdfLoadOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### TextLiteralPathDirectory
```powershell
Add-OfficePdfStamp -Text <string> -OutputDirectory <string> -LiteralPath <string[]> [-Size <OfficeLength>] [-Color <OfficeColor>] [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ImageLiteralPathDirectory
```powershell
Add-OfficePdfStamp -OutputDirectory <string> -LiteralPath <string[]> -ImagePath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ImageDocument
```powershell
Add-OfficePdfStamp -ImagePath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-PassThru] [-Document <PdfDocument>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ContentLiteralPathDirectory
```powershell
Add-OfficePdfStamp -OutputDirectory <string> -LiteralPath <string[]> -Content <scriptblock> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-ConfigureRendering <scriptblock>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ContentDocument
```powershell
Add-OfficePdfStamp [-Content] <scriptblock> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-PassThru] [-Document <PdfDocument>] [-ConfigureRendering <scriptblock>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PageLiteralPathDirectory
```powershell
Add-OfficePdfStamp -OutputDirectory <string> -LiteralPath <string[]> -OverlayPath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [-OverlayPage <int>] [-Fit <PdfPageOverlayFit>] [-HorizontalAlignment <OfficeAlignment>] [-VerticalAlignment <OfficeTextVerticalAlignment>] [-OverlayPassword <securestring>] [-IgnoreOverlayPermissionRestrictions] [-OverlayLoadOptions <PdfLoadOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PageDocument
```powershell
Add-OfficePdfStamp -OverlayPath <string> [-Page <PdfPageSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Rotation <double>] [-Opacity <double>] [-Behind] [-PassThru] [-Document <PdfDocument>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-OverlayPage <int>] [-Fit <PdfPageOverlayFit>] [-HorizontalAlignment <OfficeAlignment>] [-VerticalAlignment <OfficeTextVerticalAlignment>] [-OverlayPassword <securestring>] [-IgnoreOverlayPermissionRestrictions] [-OverlayLoadOptions <PdfLoadOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Stamps text, images, native canvas content or a PDF page onto existing PDFs.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePdfStamp -Text 'Value'
```


### EXAMPLE 2
```powershell
Add-OfficePdfStamp -Path @('C:\Path') -OutputPath 'C:\Path' -Content { }
```


### EXAMPLE 3
```powershell
Add-OfficePdfStamp -Path @('C:\Path') -OutputDirectory 'Value' -Content { }
```


## PARAMETERS

### -Behind
Place the stamp behind existing page content.

```yaml
Type: SwitchParameter
Parameter Sets: TextDocument, Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ImageDocument, ContentLiteralPathDirectory, ContentDocument, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Native Color setting for the stamp.

```yaml
Type: OfficeColor
Parameter Sets: TextDocument, Text, TextDirectory, TextLiteralPath, TextLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConfigureRendering
Configure native rendering options for the canvas callback.

```yaml
Type: ScriptBlock
Parameter Sets: Content, ContentDirectory, ContentLiteralPath, ContentLiteralPathDirectory, ContentDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Native canvas callback receiving the canvas and target page context.

```yaml
Type: ScriptBlock
Parameter Sets: Content, ContentDirectory, ContentLiteralPath, ContentLiteralPathDirectory, ContentDocument
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned PDF; omission uses the enclosing editing block.

```yaml
Type: PdfDocument
Parameter Sets: TextDocument, ImageDocument, ContentDocument, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Fit
Native Fit setting for the stamp.

```yaml
Type: PdfPageOverlayFit
Parameter Sets: Page, PageDirectory, PageLiteralPath, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values: None, Contain, Cover, Stretch

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Native Height setting for the stamp.

```yaml
Type: OfficeLength
Parameter Sets: Image, ImageDirectory, Page, PageDirectory, ImageLiteralPath, PageLiteralPath, ImageLiteralPathDirectory, ImageDocument, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HorizontalAlignment
Native HorizontalAlignment setting for the stamp.

```yaml
Type: OfficeAlignment
Parameter Sets: Page, PageDirectory, PageLiteralPath, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnoreOverlayPermissionRestrictions
After overlay authentication, explicitly ignore owner-imposed restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Page, PageDirectory, PageLiteralPath, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After authentication, explicitly ignore owner-imposed restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ContentLiteralPathDirectory, PageLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImagePath
Literal image file to stamp.

```yaml
Type: String
Parameter Sets: Image, ImageDirectory, ImageLiteralPath, ImageLiteralPathDirectory, ImageDocument
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal PDF files, including FileInfo.FullName.

```yaml
Type: String[]
Parameter Sets: TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ContentLiteralPathDirectory, PageLiteralPathDirectory
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
Native input parsing, authentication and resource policy.

```yaml
Type: PdfLoadOptions
Parameter Sets: Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ContentLiteralPathDirectory, PageLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Opacity
Native Opacity setting for the stamp.

```yaml
Type: Double
Parameter Sets: TextDocument, Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ImageDocument, ContentLiteralPathDirectory, ContentDocument, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Output folder retaining source file names.

```yaml
Type: String
Parameter Sets: TextDirectory, ImageDirectory, ContentDirectory, PageDirectory, TextLiteralPathDirectory, ImageLiteralPathDirectory, ContentLiteralPathDirectory, PageLiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Single output PDF file.

```yaml
Type: String
Parameter Sets: Text, Image, Content, Page, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverlayLoadOptions
Native loading policy for the imported overlay PDF.

```yaml
Type: PdfLoadOptions
Parameter Sets: Page, PageDirectory, PageLiteralPath, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverlayPage
One-based page imported from the overlay PDF.

```yaml
Type: Int32
Parameter Sets: Page, PageDirectory, PageLiteralPath, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverlayPassword
Password for the imported overlay PDF.

```yaml
Type: SecureString
Parameter Sets: Page, PageDirectory, PageLiteralPath, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverlayPath
Literal PDF file whose selected page is imported.

```yaml
Type: String
Parameter Sets: Page, PageDirectory, PageLiteralPath, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
Native one-based page selector, including last, odd and even.

```yaml
Type: PdfPageSelector
Parameter Sets: TextDocument, Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ImageDocument, ContentLiteralPathDirectory, ContentDocument, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo or the updated live document.

```yaml
Type: SwitchParameter
Parameter Sets: TextDocument, Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ImageDocument, ContentLiteralPathDirectory, ContentDocument, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Source PDF password.

```yaml
Type: SecureString
Parameter Sets: Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ContentLiteralPathDirectory, PageLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Input PDF files with FileSystem wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Rotation
Native Rotation setting for the stamp.

```yaml
Type: Double
Parameter Sets: TextDocument, Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ImageDocument, ContentLiteralPathDirectory, ContentDocument, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Text size; omission keeps the native 24-point default.

```yaml
Type: OfficeLength
Parameter Sets: TextDocument, Text, TextDirectory, TextLiteralPath, TextLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Literal text to stamp.

```yaml
Type: String
Parameter Sets: TextDocument, Text, TextDirectory, TextLiteralPath, TextLiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VerticalAlignment
Native VerticalAlignment setting for the stamp.

```yaml
Type: OfficeTextVerticalAlignment
Parameter Sets: Page, PageDirectory, PageLiteralPath, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values: Top, Center, Bottom

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Native Width setting for the stamp.

```yaml
Type: OfficeLength
Parameter Sets: Image, ImageDirectory, Page, PageDirectory, ImageLiteralPath, PageLiteralPath, ImageLiteralPathDirectory, ImageDocument, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Native X setting for the stamp.

```yaml
Type: OfficeLength
Parameter Sets: TextDocument, Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ImageDocument, ContentLiteralPathDirectory, ContentDocument, PageLiteralPathDirectory, PageDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Native Y setting for the stamp.

```yaml
Type: OfficeLength
Parameter Sets: TextDocument, Text, TextDirectory, Image, ImageDirectory, Content, ContentDirectory, Page, PageDirectory, TextLiteralPath, ImageLiteralPath, ContentLiteralPath, PageLiteralPath, TextLiteralPathDirectory, ImageLiteralPathDirectory, ImageDocument, ContentLiteralPathDirectory, ContentDocument, PageLiteralPathDirectory, PageDocument
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

- `System.String[]`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
