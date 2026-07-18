using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word.Pdf;
using System.Collections.Generic;
using System.Text;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using OfficeWordDocument = OfficeIMO.Word.WordDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void PdfSemanticImport_ResultsKeepDiagnosticsScopedToEachConversion() {
        byte[] imagePdf = PdfCore.PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(1, 1), 24, 24, alternativeText: "Operation-scoped image")
            .ToBytes();
        byte[] textPdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("A clean second conversion."))
            .ToBytes();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult first = LoadSemanticPdf(imagePdf).ToWordDocumentResult(options);
        using OfficeWordDocument firstDocument = first.RequireValue();
        Assert.Contains(first.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");

        PdfWordConversionResult second = LoadSemanticPdf(textPdf).ToWordDocumentResult(options);
        using OfficeWordDocument secondDocument = second.RequireValue();

        Assert.DoesNotContain(second.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.Contains(first.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
    }

    [Fact]
    public void PdfSemanticImport_SaveAsWord_ImportsEditableDocumentStructure() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                CreateOutlineFromHeadings = true,
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "PDF semantic Word import", author: "OfficeIMO", subject: "Editable import", keywords: "pdf,word,semantic")
            .H1("PDF Semantic Import", linkUri: "https://example.com/pdf-semantic-import", linkContents: "PDF Semantic Import")
            .Paragraph(paragraph => paragraph.Text("Logical paragraph survives as editable Word text."))
            .Bullets(new[] { "Editable bullet item" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(PdfPngTestImages.CreateRgbPng(1, 1), 24, 24, alternativeText: "Semantic import pixel")
            .PageBreak()
            .H1("Second Page")
            .TextField("Approval", width: 120, value: "Ready")
            .ToBytes();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        using var document = new MemoryStream();
        importedDocument.Save(document);

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfFormWidgetPlaceholder");
        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfUriLinkReconstructed");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfLinkAnnotationNotReconstructed");

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(document.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Equal("PDF semantic Word import", package.PackageProperties.Title);
        Assert.Equal("OfficeIMO", package.PackageProperties.Creator);
        Assert.Equal("Editable import", package.PackageProperties.Subject);
        Assert.Equal("pdf,word,semantic", package.PackageProperties.Keywords);

        Body body = GetPdfSemanticBody(package);
        Hyperlink headingLink = Assert.Single(body.Descendants<Hyperlink>(), link => ReadHyperlinkText(link) == "PDF Semantic Import");
        HyperlinkRelationship relationship = Assert.Single(package.MainDocumentPart!.HyperlinkRelationships, item => item.Id == headingLink.Id);
        Assert.Equal(new Uri("https://example.com/pdf-semantic-import"), relationship.Uri);
        Assert.Single(package.MainDocumentPart.ImageParts);
        Assert.NotEmpty(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.Contains(body.Descendants<Text>(), text => text.Text == "PDF Semantic Import");
        Assert.Contains(body.Descendants<Text>(), text => text.Text == "Logical paragraph survives as editable Word text.");
        Assert.Contains(body.Descendants<Text>(), text => text.Text == "Editable bullet item");
        Assert.Contains(body.Descendants<Text>(), text => text.Text == "Second Page");
        Assert.DoesNotContain(body.Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
        Assert.Contains(body.Descendants<Text>(), text => text.Text.Contains("[PDF form Tx: Approval = Ready]", StringComparison.Ordinal));
        Assert.Contains(body.Descendants<Break>(), item => item.Type?.Value == BreakValues.Page);

        Paragraph heading = Assert.Single(body.Elements<Paragraph>(), paragraph => ReadParagraphText(paragraph) == "PDF Semantic Import");
        Assert.Equal("Heading1", heading.ParagraphProperties?.ParagraphStyleId?.Val?.Value);

        Paragraph listItem = Assert.Single(body.Elements<Paragraph>(), paragraph => ReadParagraphText(paragraph) == "Editable bullet item");
        Assert.NotNull(listItem.ParagraphProperties?.NumberingProperties);

        Table table = Assert.Single(body.Descendants<Table>());
        List<TableRow> rows = table.Elements<TableRow>().ToList();
        Assert.Equal(3, rows.Count);
        Assert.NotNull(rows[0].TableRowProperties?.GetFirstChild<TableHeader>());
        Assert.Equal(new[] { "Code", "Name", "Qty" }, ReadPdfSemanticRowText(rows[0]));
        Assert.Equal(new[] { "A-100", "Alpha", "2" }, ReadPdfSemanticRowText(rows[1]));
        Assert.Equal(new[] { "B-200", "Beta", "14" }, ReadPdfSemanticRowText(rows[2]));
        Assert.Equal(JustificationValues.Right, ReadPdfSemanticCellAlignment(rows[1], 2));
    }

    [Fact]
    public void PdfSemanticImport_PageRanges_ImportsOnlySelectedPages() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("First Page Marker")
            .PageBreak()
            .H1("Second Page Marker")
            .Paragraph(paragraph => paragraph.Text("Only selected page body."))
            .ToBytes();
        var options = new PdfWordReadOptions();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.LoadPageRanges(
            pdf,
            CreateSemanticLayoutOptions(),
            new[] { PdfCore.PdfPageRange.From(2, 2) });

        PdfWordConversionResult conversion = logical.ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        string text = string.Concat(GetPdfSemanticBody(package).Descendants<Text>().Select(item => item.Text));
        Assert.DoesNotContain("First Page Marker", text, StringComparison.Ordinal);
        Assert.Contains("Second Page Marker", text, StringComparison.Ordinal);
        Assert.Contains("Only selected page body.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfSemanticImport_LoadedDocumentPageRanges_ImportsOnlySelectedPages() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("Loaded First Page Marker")
            .PageBreak()
            .H1("Loaded Second Page Marker")
            .Paragraph(paragraph => paragraph.Text("Loaded selected page body."))
            .ToBytes();
        PdfCore.PdfReadDocument readDocument = PdfCore.PdfReadDocument.Open(pdf);
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.FromPageRanges(
            readDocument,
            CreateSemanticLayoutOptions(),
            new[] { PdfCore.PdfPageRange.From(2, 2) });
        using OfficeWordDocument document = logical.ToWordDocument();

        Body body = document._wordprocessingDocument!.MainDocumentPart!.Document.Body!;
        string text = string.Concat(body.Descendants<Text>().Select(item => item.Text));
        Assert.DoesNotContain("Loaded First Page Marker", text, StringComparison.Ordinal);
        Assert.Contains("Loaded Second Page", text, StringComparison.Ordinal);
        Assert.Contains("Loaded selected page body.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfSemanticImport_UnsafeUriLinks_AreDiagnosticsNotActiveRelationships() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("Unsafe Link", linkUri: "javascript:alert(1)", linkContents: "Unsafe Link")
            .Paragraph(paragraph => paragraph.Text("The unsafe PDF action remains inert in Word."))
            .ToBytes();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfUriLinkSkippedUnsafe");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfUriLinkReconstructed");

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Empty(package.MainDocumentPart!.HyperlinkRelationships);
        Assert.Empty(GetPdfSemanticBody(package).Descendants<Hyperlink>());
    }

    [Fact]
    public void PdfSemanticImport_DisabledUriLinks_DoNotCreateActiveWordHyperlinks() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("Disabled URI Link", linkUri: "https://example.com/disabled", linkContents: "Disabled URI Link")
            .ToBytes();
        var options = new PdfWordReadOptions {
            ImportUriLinks = false,
            ImportInternalLinks = true,

        };

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfUriLinkReconstructed");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Empty(package.MainDocumentPart!.HyperlinkRelationships);
        Assert.Empty(GetPdfSemanticBody(package).Descendants<Hyperlink>());
        Assert.Contains(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text == "Disabled URI Link");
    }

    [Fact]
    public void PdfSemanticImport_InternalLinks_BecomeWordBookmarkHyperlinks() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                CreateOutlineFromHeadings = true
            })
            .H1("Jump to details", linkDestinationName: "Details", linkContents: "Heading jump metadata")
            .Paragraph(paragraph => paragraph.Text("Introductory text before the destination."))
            .Bookmark("Details")
            .H2("Details")
            .Paragraph(paragraph => paragraph.Text("Destination content survives as editable Word text."))
            .ToBytes();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfInternalLinkReconstructed");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfLinkAnnotationNotReconstructed");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Empty(package.MainDocumentPart!.HyperlinkRelationships);

        Body body = GetPdfSemanticBody(package);
        Hyperlink hyperlink = Assert.Single(body.Descendants<Hyperlink>(), link => ReadHyperlinkText(link) == "Jump to details");
        string anchor = Assert.IsType<string>(hyperlink.Anchor?.Value);
        Assert.StartsWith("OfficeIMO_Pdf_Dest_Details", anchor, StringComparison.Ordinal);
        Assert.Contains(body.Descendants<BookmarkStart>(), bookmark => bookmark.Name?.Value == anchor);
        Assert.Contains(body.Descendants<Text>(), text => text.Text == "Details");
    }

    [Fact]
    public void PdfSemanticImport_DisabledImageImport_UsesEditablePlaceholder() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("Image Placeholder Policy")
            .Image(PdfPngTestImages.CreateRgbPng(1, 1), 24, 24, alternativeText: "Placeholder policy pixel")
            .ToBytes();
        var options = new PdfWordReadOptions {
            ImportImages = false,
            IncludeImagePlaceholders = true,

        };

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Empty(package.MainDocumentPart!.ImageParts);
        Assert.Contains(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
        Assert.Contains(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("image-import-disabled", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_RawDeviceRgbImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildRawDeviceRgbImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_FilterChainDeviceRgbImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildAsciiHexFlateDeviceRgbImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_DctImageStreamsWithSoftMask_AreEmbeddedWithTransparencyWarning() {
        byte[] pdf = BuildDeviceRgbJpegSoftMaskImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.Contains(conversion.Report.Warnings, warning =>
            warning.Code == "PdfImageTransparencyMaskNotResolved" &&
            warning.Details.TryGetValue("MaskKind", out string? maskKind) &&
            maskKind == "soft-mask");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_ColorKeyMaskedImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildDeviceRgbColorKeyMaskImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_DeviceCmykImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildRawDeviceCmykImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_DeviceCmykSoftMaskImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildDeviceCmykSoftMaskImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_IndexedImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildRawIndexedRgbImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_IndexedSoftMaskImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildIndexedSoftMaskImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_IndexedColorKeyMaskedImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildIndexedColorKeyMaskImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_ImageMaskStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildImageMaskPdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_DecodeRemappedImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildRawDeviceGrayInvertedDecodeImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfSemanticImport_IccBasedImageStreams_AreEmbeddedAsNativeWordImages() {
        byte[] pdf = BuildRawIccBasedRgbImagePdf();
        var options = new PdfWordReadOptions();

        PdfWordConversionResult conversion = LoadSemanticPdf(pdf).ToWordDocumentResult(options);
        using OfficeWordDocument importedDocument = conversion.Value;
        byte[] documentBytes = importedDocument.ToBytes();

        Assert.Contains(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "PdfImageEmbeddingSkipped");
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.NotEmpty(GetPdfSemanticBody(package).Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.DoesNotContain(GetPdfSemanticBody(package).Descendants<Text>(), text => text.Text.Contains("[PDF image: page 1", StringComparison.Ordinal));
    }

    private static PdfCore.PdfLogicalDocument LoadSemanticPdf(byte[] pdf) =>
        PdfCore.PdfLogicalDocument.Load(pdf, CreateSemanticLayoutOptions());

    private static PdfCore.PdfTextLayoutOptions CreateSemanticLayoutOptions() => new PdfCore.PdfTextLayoutOptions {
        ForceSingleColumn = true
    };

    private static Body GetPdfSemanticBody(WordprocessingDocument package) {
        return package.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("The saved document body is missing.");
    }

    private static string ReadParagraphText(Paragraph paragraph) {
        return string.Concat(paragraph.Descendants<Text>().Select(text => text.Text ?? string.Empty));
    }

    private static string ReadHyperlinkText(Hyperlink hyperlink) {
        return string.Concat(hyperlink.Descendants<Text>().Select(text => text.Text ?? string.Empty));
    }

    private static string[] ReadPdfSemanticRowText(TableRow row) {
        return row.Elements<TableCell>()
            .Select(cell => string.Concat(cell.Descendants<Text>().Select(text => text.Text ?? string.Empty)))
            .ToArray();
    }

    private static JustificationValues? ReadPdfSemanticCellAlignment(TableRow row, int columnIndex) {
        return row.Elements<TableCell>()
            .ElementAt(columnIndex)
            .Elements<Paragraph>()
            .FirstOrDefault()?
            .ParagraphProperties?
            .Justification?
            .Val?
            .Value;
    }

    private static byte[] BuildRawDeviceRgbImagePdf() {
        return BuildDeviceRgbImagePdf("abc", string.Empty);
    }

    private static byte[] BuildRawDeviceCmykImagePdf() {
        return BuildImagePdf("DeviceCMYK", "abc ", string.Empty);
    }

    private static byte[] BuildDeviceCmykSoftMaskImagePdf() {
        return BuildImagePdfWithColorSpace(
            "/DeviceCMYK",
            1,
            1,
            8,
            "abc ",
            " /SMask 6 0 R",
            new[] { BuildSoftMaskObject(6, 126) },
            7);
    }

    private static byte[] BuildDeviceRgbJpegSoftMaskImagePdf() {
        byte[] jpeg = CreateMinimalJpeg(2, 1);
        return BuildImagePdfWithColorSpace(
            "/DeviceRGB",
            2,
            1,
            8,
            PdfCore.PdfEncoding.Latin1GetString(jpeg),
            " /Filter /DCTDecode /SMask 6 0 R",
            new[] { BuildSoftMaskObject(6, new byte[] { 126, 64 }, 2, 1) },
            7);
    }

    private static byte[] BuildRawDeviceGrayInvertedDecodeImagePdf() {
        return BuildImagePdf("DeviceGray", "a", " /Decode [1 0]");
    }

    private static byte[] BuildRawIccBasedRgbImagePdf() {
        return BuildIccBasedImagePdf(3, "DeviceRGB", "abc");
    }

    private static byte[] BuildRawIndexedRgbImagePdf() {
        return BuildImagePdfWithColorSpace("[/Indexed /DeviceRGB 1 <FF000000FF00>]", 2, 1, 1, "@", string.Empty);
    }

    private static byte[] BuildIndexedSoftMaskImagePdf() {
        return BuildImagePdfWithColorSpace(
            "[/Indexed /DeviceRGB 1 <FF000000FF00>]",
            2,
            1,
            1,
            "@",
            " /SMask 6 0 R",
            new[] { BuildSoftMaskObject(6, new byte[] { 126, 64 }, 2, 1) },
            7);
    }

    private static byte[] BuildIndexedColorKeyMaskImagePdf() {
        return BuildImagePdfWithColorSpace(
            "[/Indexed /DeviceRGB 1 <FF000000FF00>]",
            2,
            1,
            1,
            "@",
            " /Mask [0 0]");
    }

    private static byte[] BuildImageMaskPdf() {
        string content = string.Join("\n", new[] {
            "q",
            "24 0 0 24 36 160 cm",
            "/Im1 Do",
            "Q"
        });

        string imageStreamData = "@";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 220 220] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /ImageMask true /Length " + Encoding.ASCII.GetByteCount(imageStreamData).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            imageStreamData,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildAsciiHexFlateDeviceRgbImagePdf() {
        string encoded = EncodeAsciiHex(BuildStoredZlib(new byte[] { 97, 98, 99 }));
        return BuildDeviceRgbImagePdf(encoded, " /Filter [/ASCIIHexDecode /FlateDecode]");
    }

    private static byte[] BuildDeviceRgbColorKeyMaskImagePdf() {
        return BuildImagePdfWithColorSpace("/DeviceRGB", 2, 1, 8, "ABCDEF", " /Mask [97 97 98 98 99 99]");
    }

    private static byte[] BuildDeviceRgbImagePdf(string imageStreamData, string imageDictionarySuffix) {
        return BuildImagePdf("DeviceRGB", imageStreamData, imageDictionarySuffix);
    }

    private static byte[] BuildImagePdf(string colorSpace, string imageStreamData, string imageDictionarySuffix) {
        return BuildImagePdfWithColorSpace("/" + colorSpace, 1, 1, 8, imageStreamData, imageDictionarySuffix);
    }

    private static byte[] BuildImagePdfWithColorSpace(
        string colorSpaceObject,
        int width,
        int height,
        int bitsPerComponent,
        string imageStreamData,
        string imageDictionarySuffix) {
        return BuildImagePdfWithColorSpace(colorSpaceObject, width, height, bitsPerComponent, imageStreamData, imageDictionarySuffix, Array.Empty<string>(), 6);
    }

    private static byte[] BuildIccBasedImagePdf(int componentCount, string alternateColorSpace, string imageStreamData) {
        string profileObject = string.Join("\n", new[] {
            "6 0 obj",
            "<< /N " + componentCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Alternate /" + alternateColorSpace + " /Length 0 >>",
            "stream",
            string.Empty,
            "endstream",
            "endobj"
        });
        return BuildImagePdfWithColorSpace("[/ICCBased 6 0 R]", 1, 1, 8, imageStreamData, string.Empty, new[] { profileObject }, 7);
    }

    private static string BuildSoftMaskObject(int objectNumber, byte alpha) {
        return BuildSoftMaskObject(objectNumber, new[] { alpha }, 1, 1);
    }

    private static string BuildSoftMaskObject(int objectNumber, byte[] alpha, int width, int height) {
        string encodedAlpha = EncodeAsciiHex(BuildStoredZlib(alpha));
        return string.Join("\n", new[] {
            objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj",
            "<< /Type /XObject /Subtype /Image /Width "
                + width.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /Height "
                + height.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter [/ASCIIHexDecode /FlateDecode] /Length "
                + Encoding.ASCII.GetByteCount(encodedAlpha).ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " >>",
            "stream",
            encodedAlpha,
            "endstream",
            "endobj"
        });
    }

    private static byte[] BuildImagePdfWithColorSpace(
        string colorSpaceObject,
        int width,
        int height,
        int bitsPerComponent,
        string imageStreamData,
        string imageDictionarySuffix,
        IReadOnlyList<string> additionalObjects,
        int trailerSize) {
        string content = string.Join("\n", new[] {
            "q",
            "24 0 0 24 36 160 cm",
            "/Im1 Do",
            "Q"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 220 220] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width "
                + width.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /Height "
                + height.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /ColorSpace "
                + colorSpaceObject
                + " /BitsPerComponent "
                + bitsPerComponent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /Length "
                + imageStreamData.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + imageDictionarySuffix
                + " >>",
            "stream",
            imageStreamData,
            "endstream",
            "endobj",
            string.Join("\n", additionalObjects),
            "trailer",
            "<< /Root 1 0 R /Size " + trailerSize.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "%%EOF"
        }) + "\n";

        return PdfCore.PdfEncoding.Latin1GetBytes(pdf);
    }

    private static byte[] CreateMinimalJpeg(int width, int height) {
        return new byte[] {
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            (byte)(height >> 8), (byte)(height & 0xFF),
            (byte)(width >> 8), (byte)(width & 0xFF),
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        };
    }

    private static byte[] BuildStoredZlib(byte[] data) {
        using var ms = new MemoryStream();
        ms.WriteByte(0x78);
        ms.WriteByte(0x01);
        ms.WriteByte(0x01);
        ms.WriteByte((byte)(data.Length & 0xFF));
        ms.WriteByte((byte)((data.Length >> 8) & 0xFF));
        int nlen = data.Length ^ 0xFFFF;
        ms.WriteByte((byte)(nlen & 0xFF));
        ms.WriteByte((byte)((nlen >> 8) & 0xFF));
        ms.Write(data, 0, data.Length);
        uint adler = Adler32(data);
        ms.WriteByte((byte)((adler >> 24) & 0xFF));
        ms.WriteByte((byte)((adler >> 16) & 0xFF));
        ms.WriteByte((byte)((adler >> 8) & 0xFF));
        ms.WriteByte((byte)(adler & 0xFF));
        return ms.ToArray();
    }

    private static string EncodeAsciiHex(byte[] data) {
        var builder = new StringBuilder(data.Length * 2 + 1);
        for (int i = 0; i < data.Length; i++) {
            builder.Append(data[i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }

        builder.Append('>');
        return builder.ToString();
    }

    private static uint Adler32(byte[] data) {
        const uint mod = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }

        return (b << 16) | a;
    }
}
