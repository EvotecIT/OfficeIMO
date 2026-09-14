using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfHtmlPageSelectionTests {
    [Fact]
    public void DocumentHtmlExport_AppliesPageRangesBeforeSemanticReading() {
        PdfDocument source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page 1"));
        for (int pageNumber = 2; pageNumber <= 1001; pageNumber++) {
            int capturedPageNumber = pageNumber;
            source.PageBreak().Paragraph(paragraph => paragraph.Text("Page " + capturedPageNumber));
        }
        PdfDocument document = PdfDocument.Load(source.ToBytes());
        var options = new PdfToHtmlOptions {
            ReadOptions = new PdfReadOptions {
                Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 1 }
            },
            PageRanges = new[] { PdfPageRange.From(1001, 1001) }
        };

        PdfHtmlConversionResult result = document.ToHtmlResult(options);

        Assert.Contains("Page 1001", result.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("Page 1<", result.Value, StringComparison.Ordinal);
        Assert.Equal(1001, result.Summary.SourcePageCount);
        Assert.Equal(1, result.Summary.RenderedPageCount);
        Assert.Equal(new[] { 1001 }, result.Summary.PageNumbers);
    }

    [Fact]
    public void HtmlExport_DoesNotReportUnusedOptionalContentDefinitionsAsLoss() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(
            PdfOptionalContentSupport.BuildOptionalContentMetadataPdf());

        PdfHtmlConversionResult result = logical.ToHtmlResult();
        PdfTableExtractionScopeReport tableScope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument wordDocument = word.Value;

        Assert.True(logical.HasOptionalContentGroups);
        Assert.All(logical.Pages, static page => Assert.False(page.HasOptionalContentUsage));
        Assert.Equal(logical.OptionalContentGroupCount, tableScope.OptionalContentGroupCount);
        Assert.Equal(0, tableScope.PagesWithOptionalContent);
        Assert.DoesNotContain(word.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.DoesNotContain(result.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
    }

    [Fact]
    public void HtmlExport_ReportsOptionalContentOnlyWhenSelectedPagesUseIt() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Unlayered page"))
            .PageBreak()
            .Layer("Selected layer", layer => layer.Paragraph(paragraph => paragraph.Text("Layered page")))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(source);

        PdfHtmlConversionResult firstPage = document.ToHtmlResult(new PdfToHtmlOptions {
            PageRanges = new[] { PdfPageRange.From(1, 1) }
        });
        PdfHtmlConversionResult secondPage = document.ToHtmlResult(new PdfToHtmlOptions {
            PageRanges = new[] { PdfPageRange.From(2, 2) }
        });
        PdfDocumentReadResult firstPageLogical = document.Read(new PdfReadOptions {
            PageSelection = PdfPageSelection.From(1)
        });
        PdfDocumentReadResult secondPageLogical = document.Read(new PdfReadOptions {
            PageSelection = PdfPageSelection.From(2)
        });
        PdfWordConversionResult firstPageWord = firstPageLogical.ToWordDocumentResult();
        PdfWordConversionResult secondPageWord = secondPageLogical.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument firstWordDocument = firstPageWord.Value;
        using OfficeIMO.Word.WordDocument secondWordDocument = secondPageWord.Value;

        Assert.DoesNotContain(firstPage.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.Contains(secondPage.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened" &&
            warning.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Equal(0, PdfLogicalTableAnalysis.AnalyzeExtractionScope(firstPageLogical).PagesWithOptionalContent);
        Assert.Equal(1, PdfLogicalTableAnalysis.AnalyzeExtractionScope(secondPageLogical).PagesWithOptionalContent);
        Assert.DoesNotContain(firstPageWord.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.Contains(secondPageWord.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
    }
}
