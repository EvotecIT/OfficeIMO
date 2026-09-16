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
    }
}
