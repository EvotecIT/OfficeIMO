using System.Text;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReverseAttachmentLossTests {
    [Fact]
    public void SelectedPageWordConversionReportsSourceTaggedStructureLoss() {
        byte[] source = PdfDocument.Create(new PdfOptions {
                TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
            })
            .Paragraph(paragraph => paragraph.Text("First invoice page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second invoice page"))
            .ToBytes();
        PdfDocument opened = PdfDocument.Load(source);
        PdfDocumentReadResult selected = opened.Read(new PdfReadOptions {
            PageSelection = PdfPageSelection.From(1)
        });
        Assert.Null(selected.TaggedContent);

        PdfWordConversionResult word = opened.ToWordDocumentResult(new PdfToWordOptions {
            ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(1) }
        });
        using (word.Value) {
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfTaggedStructureNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission &&
                int.Parse(warning.Details["StructureElementCount"], System.Globalization.CultureInfo.InvariantCulture) > 0);
        }

        PdfHtmlConversionResult html = opened.ToHtmlResult(new PdfToHtmlOptions {
            PageRanges = new[] { PdfPageRange.From(1, 1) }
        });
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "PdfTaggedStructureOmitted" &&
            warning.LossKind == OfficeConversionLossKind.Omission);

        PdfToWordOptions visualWordOptions = PdfToWordOptions.CreateVisualPages();
        visualWordOptions.ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(1) };
        PdfWordConversionResult visualWord = opened.ToWordDocumentResult(visualWordOptions);
        using (visualWord.Value) {
            Assert.True(visualWord.HasLoss);
            Assert.Contains(visualWord.Report.Warnings, static warning =>
                warning.Code == "PdfTaggedStructureNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }

        PdfToPowerPointOptions visualPowerPointOptions = PdfToPowerPointOptions.CreateVisualPages();
        visualPowerPointOptions.ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(1) };
        PdfPowerPointConversionResult visualPowerPoint = opened.ToPowerPointPresentationResult(visualPowerPointOptions);
        using (visualPowerPoint.Value) {
            Assert.True(visualPowerPoint.Report.HasOmittedPageContent);
            Assert.Contains(visualPowerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfTaggedStructureNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Fact]
    public void AttachmentOnlyPdfReportsLossInEditableWordAndHtml() {
        byte[] source = PdfDocument.Create()
            .AttachFile("invoice.xml", Encoding.UTF8.GetBytes("<invoice />"), "application/xml")
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        Assert.Single(logical.Attachments);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.True(word.HasLoss);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfAttachmentsNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission &&
                warning.Details["AttachmentCount"] == "1");
            Assert.Throws<InvalidOperationException>(() => word.RequireNoLoss());
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.True(html.HasLoss);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "PdfAttachmentsOmitted" &&
            warning.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<InvalidOperationException>(() => html.RequireNoLoss());

        byte[] visualSource = PdfDocument.Create()
            .AttachFile("invoice.xml", Encoding.UTF8.GetBytes("<invoice />"), "application/xml")
            .Paragraph(paragraph => paragraph.Text("Invoice"))
            .ToBytes();
        PdfWordConversionResult visualWord = PdfDocument.Load(visualSource)
            .ToWordDocumentResult(PdfToWordOptions.CreateVisualPages());
        using (visualWord.Value) {
            Assert.Contains(visualWord.Report.Warnings, static warning =>
                warning.Code == "PdfAttachmentsNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission &&
                warning.Details["AttachmentCount"] == "1");
        }
    }

    [Fact]
    public void SelectedPageConversionsStillReportDocumentWideAttachmentLoss() {
        byte[] source = PdfDocument.Create()
            .AttachFile("invoice.xml", Encoding.UTF8.GetBytes("<invoice />"), "application/xml")
            .Paragraph(paragraph => paragraph.Text("First invoice page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second invoice page"))
            .ToBytes();
        PdfDocument opened = PdfDocument.Load(source);
        PdfDocumentReadResult selected = opened.Read(new PdfReadOptions {
            PageSelection = PdfPageSelection.From(1)
        });
        Assert.Empty(selected.Attachments);
        PdfTableExtractionScopeReport scope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(selected);
        Assert.Equal(1, scope.AttachmentCount);

        PdfWordConversionResult word = opened.ToWordDocumentResult(new PdfToWordOptions {
            ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(1) }
        });
        using (word.Value) {
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfAttachmentsNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission &&
                warning.Details["AttachmentCount"] == "1");
        }

        PdfHtmlConversionResult html = opened.ToHtmlResult(new PdfToHtmlOptions {
            PageRanges = new[] { PdfPageRange.From(1, 1) }
        });
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "PdfAttachmentsOmitted" &&
            warning.LossKind == OfficeConversionLossKind.Omission);

        PdfExcelTableImportResult excel = selected.ImportTablesToExcelDocumentResult();
        PdfPowerPointConversionResult powerPoint = selected.ToPowerPointPresentationResult();
        using (excel.Value)
        using (powerPoint.Value) {
            Assert.Equal(1, excel.Report.SourceScope.AttachmentCount);
            Assert.Equal(1, powerPoint.Report.SourceScope.AttachmentCount);
            Assert.True(excel.HasLoss);
            Assert.True(powerPoint.HasLoss);
        }

        PdfToPowerPointOptions visualOptions = PdfToPowerPointOptions.CreateVisualPages();
        visualOptions.ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(1) };
        PdfPowerPointConversionResult visual = opened.ToPowerPointPresentationResult(visualOptions);
        using (visual.Value) {
            Assert.True(visual.Report.HasOmittedPageContent);
            Assert.Contains(visual.Report.Warnings, static warning =>
                warning.Code == "PdfAttachmentsNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission &&
                warning.Details["Count"] == "1");
            Assert.Throws<InvalidOperationException>(() => visual.RequireNoLoss());
        }
    }

    [Fact]
    public void VisualWordAndPowerPointReportDocumentNavigationOmissions() {
        PdfOptions options = new PdfOptions { CreateOutlineFromHeadings = true }
            .SetOpenAction(1, destinationMode: PdfOpenActionDestinationMode.Fit);
        byte[] source = PdfDocument.Create(options)
            .H1("Invoice")
            .ToBytes();
        PdfDocument opened = PdfDocument.Load(source);

        PdfWordConversionResult word = opened.ToWordDocumentResult(PdfToWordOptions.CreateVisualPages());
        using (word.Value) {
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfCatalogActionsNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfOutlineHierarchyNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }

        PdfPowerPointConversionResult powerPoint = opened.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateVisualPages());
        using (powerPoint.Value) {
            Assert.True(powerPoint.Report.HasOmittedPageContent);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfDocumentActionsNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfOutlinesNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Fact]
    public void VisualPageReportsOnlySelectedPageInteractionsAsOmitted() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Link("First", "https://example.com/first"))
            .TextAnnotation("First note")
            .PageBreak()
            .Paragraph(paragraph => paragraph.Link("Second", "https://example.com/second"))
            .TextAnnotation("Second note")
            .ToBytes();
        PdfDocument opened = PdfDocument.Load(source);
        Assert.All(opened.Inspect().Pages, static page => {
            Assert.NotEmpty(page.LinkAnnotations);
            Assert.NotEmpty(page.Annotations);
        });

        PdfToWordOptions wordOptions = PdfToWordOptions.CreateVisualPages();
        wordOptions.ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(2) };
        PdfWordConversionResult word = opened.ToWordDocumentResult(wordOptions);
        using (word.Value) {
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfLinksNotReconstructed" && warning.Source == "Page 2/Links" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfAnnotationsNotReconstructed" && warning.Source == "Page 2/Annotations" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.DoesNotContain(word.Report.Warnings, static warning => warning.Source.StartsWith("Page 1/", StringComparison.Ordinal));
        }

        PdfToPowerPointOptions powerPointOptions = PdfToPowerPointOptions.CreateVisualPages();
        powerPointOptions.ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(2) };
        PdfPowerPointConversionResult powerPoint = opened.ToPowerPointPresentationResult(powerPointOptions);
        using (powerPoint.Value) {
            Assert.True(powerPoint.Report.HasOmittedPageContent);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfLinksNotReconstructed" && warning.Source == "PDF page 2/Links" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfAnnotationsNotReconstructed" && warning.Source == "PDF page 2/Annotations" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.DoesNotContain(powerPoint.Report.Warnings, static warning => warning.Source.StartsWith("PDF page 1/", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void VisualPageReportsUnsupportedLinkActionThatTypedLinkReaderSkips() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Link("Invoice link", "https://example.com/invoice"))
            .ToBytes();
        byte[] marker = Encoding.ASCII.GetBytes("/S /URI");
        int actionOffset = -1;
        for (int index = 0; index <= source.Length - marker.Length; index++) {
            if (!source.AsSpan(index, marker.Length).SequenceEqual(marker)) continue;
            actionOffset = index;
            break;
        }
        Assert.True(actionOffset >= 0);
        source[actionOffset + 4] = (byte)'F';
        source[actionOffset + 5] = (byte)'o';
        source[actionOffset + 6] = (byte)'o';
        PdfDocument opened = PdfDocument.Load(source);
        PdfPageInfo page = Assert.Single(opened.Inspect().Pages);
        Assert.Empty(page.LinkAnnotations);
        Assert.Contains(page.Annotations, static annotation => annotation.Subtype == "Link");

        PdfWordConversionResult word = opened.ToWordDocumentResult(PdfToWordOptions.CreateVisualPages());
        using (word.Value) {
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfLinksNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
        PdfPowerPointConversionResult powerPoint = opened.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateVisualPages());
        using (powerPoint.Value) {
            Assert.True(powerPoint.Report.HasOmittedPageContent);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfLinksNotReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Theory]
    [InlineData(PdfHtmlProfile.Semantic, true, true, 0)]
    [InlineData(PdfHtmlProfile.Semantic, true, false, 0)]
    [InlineData(PdfHtmlProfile.Semantic, false, true, 3)]
    [InlineData(PdfHtmlProfile.Semantic, false, false, 4)]
    [InlineData(PdfHtmlProfile.PositionedReview, true, true, 0)]
    [InlineData(PdfHtmlProfile.PositionedReview, true, false, 4)]
    [InlineData(PdfHtmlProfile.PositionedReview, false, true, 3)]
    [InlineData(PdfHtmlProfile.PositionedReview, false, false, 4)]
    public void HtmlReportsOnlyMetadataFieldsActuallyOmitted(
        PdfHtmlProfile profile,
        bool includeMetadata,
        bool emitDocumentShell,
        int omittedFieldCount) {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Invoice"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(source).UpdateMetadata(
            title: "Invoice title", author: "Invoice author", subject: "Invoice subject", keywords: "invoice");
        var options = new PdfToHtmlOptions {
            Profile = profile,
            IncludeMetadata = includeMetadata,
            EmitDocumentShell = emitDocumentShell
        };

        PdfHtmlConversionResult result = document.ToHtmlResult(options);
        PdfConversionWarning[] warnings = result.Report.Warnings
            .Where(static item => item.Code == "PdfMetadataOmitted").ToArray();
        if (omittedFieldCount == 0) {
            Assert.Empty(warnings);
        } else {
            PdfConversionWarning warning = Assert.Single(warnings);
            Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
            Assert.StartsWith(omittedFieldCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + " PDF ",
                warning.Message, StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData(PdfHtmlProfile.Semantic)]
    [InlineData(PdfHtmlProfile.PositionedReview)]
    public void HtmlRequireNoLossRejectsDisabledMetadata(PdfHtmlProfile profile) {
        PdfDocument source = PdfDocument.Load(PdfDocument.Create()
                .Paragraph(paragraph => paragraph.Text("Invoice"))
                .ToBytes())
            .UpdateMetadata(author: "Invoice author");
        PdfHtmlConversionResult retained = source.ToHtmlResult(new PdfToHtmlOptions {
            Profile = profile,
            IncludeMetadata = true
        });
        Assert.DoesNotContain(retained.Report.Warnings, static warning => warning.Code == "PdfMetadataOmitted");

        PdfHtmlConversionResult omitted = source.ToHtmlResult(new PdfToHtmlOptions {
            Profile = profile,
            IncludeMetadata = false
        });
        Assert.True(omitted.HasLoss);
        Assert.Contains(omitted.Report.Warnings, static warning =>
            warning.Code == "PdfMetadataOmitted" && warning.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<InvalidOperationException>(() => omitted.RequireNoLoss());
    }
}
