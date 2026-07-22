using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlRichGenericImports {
    private const string PixelPng = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/69DjmQAAAABJRU5ErkJggg==";

    [Fact]
    public void ExcelHtml_GenericImportPreservesRichRunsStylesAndLinks() {
        const string html = """
            <section>
              <h1>Summary</h1>
              <p style="color:#123456;background-color:#abcdef">
                Plain <strong style="font-family:Arial;font-size:20px">bold</strong>
                <a href="https://example.com/report"><em>linked</em></a>
              </p>
            </section>
            """;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;
        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        ExcelCell richCell = Enumerable.Range(1, 8)
            .Select(row => sheet.CellAt(row, 1))
            .First(cell => cell.GetRichText().Any(run => run.Text.Contains("bold", StringComparison.Ordinal)));
        IReadOnlyList<ExcelRichTextRun> runs = richCell.GetRichText();

        ExcelRichTextRun bold = Assert.Single(runs, run => run.Text.Contains("bold", StringComparison.Ordinal));
        ExcelRichTextRun linked = Assert.Single(runs, run => run.Text.Contains("linked", StringComparison.Ordinal));
        Assert.True(bold.Bold);
        Assert.Equal("Arial", bold.FontName);
        Assert.Equal(15D, bold.FontSize);
        Assert.True(linked.Italic);
        Assert.Contains(sheet.GetHyperlinks().Values,
            hyperlink => hyperlink.Target == "https://example.com/report");
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void GenericScreenAdaptersBuildSemanticsInTheirTargetMediaContext() {
        const string html = """
            <style>
              @media print { p { font-weight: normal; } }
              @media screen { p { font-weight: bold; } }
            </style>
            <p>Screen target</p>
            """;
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Profile = HtmlConversionProfile.HighFidelityPrint
        });

        HtmlToExcelResult excelResult = source.ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = excelResult.Value;
        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        ExcelCell excelCell = Assert.Single(Enumerable.Range(1, 4)
            .Select(row => sheet.CellAt(row, 1)), cell => cell.GetValue<string>() == "Screen target");
        Assert.True(Assert.Single(excelCell.GetRichText()).Bold);

        HtmlToPowerPointResult powerPointResult = source.ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = powerPointResult.Value;
        PowerPointTextBox textBox = Assert.Single(Assert.Single(presentation.Slides).TextBoxes,
            candidate => candidate.Text == "Screen target");
        PowerPointTextRun powerPointRun = Assert.Single(Assert.Single(textBox.Paragraphs).Runs);
        Assert.True(powerPointRun.Bold);
    }

    [Fact]
    public void GenericNativeAdaptersPreserveMixedInlineContainerAsOneLinkedBlock() {
        const string html = "<div>Read <a href='https://example.test/report'>the report</a> now.</div>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlSemanticBlock block = Assert.Single(Assert.Single(source.SemanticDocument.Sections).Blocks);

        Assert.Equal(HtmlSemanticBlockKind.Paragraph, block.Kind);
        Assert.Equal("Read the report now.", block.Text);
        Assert.Equal("https://example.test/report", Assert.Single(block.Runs,
            run => run.Text.Contains("the report", StringComparison.Ordinal)).Hyperlink);

        HtmlToExcelResult excelResult = source.ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = excelResult.Value;
        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        Assert.Contains(Enumerable.Range(1, 8).Select(row => sheet.CellAt(row, 1).GetValue<string>()),
            value => value == "Read the report now.");
        Assert.Contains(sheet.GetHyperlinks().Values,
            hyperlink => hyperlink.Target == "https://example.test/report");

        HtmlToPowerPointResult powerPointResult = source.ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = powerPointResult.Value;
        PowerPointTextBox textBox = Assert.Single(Assert.Single(presentation.Slides).TextBoxes,
            candidate => candidate.Text == "Read the report now.");
        PowerPointTextRun powerPointLink = Assert.Single(textBox.Paragraphs.SelectMany(paragraph => paragraph.Runs),
            run => run.Text.Contains("the report", StringComparison.Ordinal));
        Assert.Equal(new Uri("https://example.test/report"), powerPointLink.Hyperlink);

        HtmlToOneNoteSectionResult oneNoteResult = source.ToOneNoteSectionResult();
        OneNoteParagraph paragraph = Assert.Single(Assert.Single(Assert.Single(oneNoteResult.Value.Pages).Outlines)
            .Children.OfType<OneNoteParagraph>());
        Assert.Equal("Read the report now.", string.Concat(paragraph.Runs.Select(run => run.Text)));
        Assert.Equal("https://example.test/report", Assert.Single(paragraph.Runs,
            run => run.Text.Contains("the report", StringComparison.Ordinal)).Hyperlink);
    }

    [Fact]
    public void ExcelHtml_GenericTableImportPreservesRichCellsLinksAndEmbeddedImages() {
        string html = $$"""
            <section><h1>Data</h1>
              <table><caption>Results</caption>
                <tr><th style="background-color:#abcdef;color:#123456">Name</th></tr>
                <tr><td><strong>Bold</strong> <a href="https://example.com/item"><em>linked</em></a></td></tr>
              </table>
              <img src="{{PixelPng}}" alt="Evidence" width="24" height="30">
            </section>
            """;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;
        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        IReadOnlyList<ExcelRichTextRun> runs = sheet.CellAt(2, 1).GetRichText();
        ExcelImage image = Assert.Single(sheet.Images);

        Assert.True(Assert.Single(runs, run => run.Text.Contains("Bold", StringComparison.Ordinal)).Bold);
        Assert.True(Assert.Single(runs, run => run.Text.Contains("linked", StringComparison.Ordinal)).Italic);
        Assert.Contains(sheet.GetHyperlinks().Values,
            hyperlink => hyperlink.Target == "https://example.com/item");
        Assert.Equal("Evidence", image.Description);
        Assert.Equal(24, image.WidthPixels);
        Assert.Equal(30, image.HeightPixels);
        Assert.Equal(1, result.Images);
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void ExcelHtml_GenericImportPreservesRepeatedImageOccurrences() {
        string html = "<img src='" + PixelPng + "' alt='First'><img src='" + PixelPng + "' alt='Second'>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        Assert.Single(source.SemanticDocument.Resources);
        Assert.Equal(2, source.SemanticDocument.ResourceOccurrences.Count);

        HtmlToExcelResult result = source.ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;
        ExcelImage[] images = Assert.Single(workbook.Sheets).Images.ToArray();

        Assert.Equal(2, result.Images);
        Assert.Equal(new[] { "First", "Second" }, images.Select(image => image.Description));
    }

    [Fact]
    public void PowerPointHtml_GenericImportPreservesRichRunsLinksAndNestedLists() {
        const string html = """
            <section>
              <h1>Summary</h1>
              <p>Plain <strong style="color:#123456;font-family:Arial;font-size:20px">bold</strong>
                 <a href="https://example.com/report"><em>linked</em></a></p>
              <ol><li>First<ul><li><u>Nested</u></li></ul></li><li>Second</li></ol>
            </section>
            """;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);
        PowerPointTextRun[] runs = slide.TextBoxes
            .SelectMany(textBox => textBox.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .ToArray();
        PowerPointTextRun bold = Assert.Single(runs, run => run.Text.Contains("bold", StringComparison.Ordinal));
        PowerPointTextRun linked = Assert.Single(runs, run => run.Text.Contains("linked", StringComparison.Ordinal));
        PowerPointTextBox list = Assert.Single(slide.TextBoxes,
            textBox => textBox.Paragraphs.Any(paragraph => paragraph.IsNumbered));

        Assert.True(bold.Bold);
        Assert.Equal("123456", bold.Color);
        Assert.Equal("Arial", bold.FontName);
        Assert.Equal(15, bold.FontSize);
        Assert.True(linked.Italic);
        Assert.Equal(new Uri("https://example.com/report"), linked.Hyperlink);
        Assert.True(list.Paragraphs[0].IsNumbered);
        Assert.Equal(0, list.Paragraphs[0].Level);
        Assert.Equal("•", list.Paragraphs[1].BulletCharacter);
        Assert.Equal(1, list.Paragraphs[1].Level);
        Assert.True(list.Paragraphs[1].Runs.Single().Underline);
        Assert.True(list.Paragraphs[2].IsNumbered);
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void PowerPointHtml_GenericTableImportPreservesRichRunsLinksAndCellStyles() {
        const string html = """
            <section><h1>Data</h1>
              <table>
                <tr><th style="background-color:#abcdef;color:#123456;text-align:center">Name</th></tr>
                <tr><td><strong>Bold</strong> <a href="https://example.com/item"><em>linked</em></a></td></tr>
              </table>
            </section>
            """;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;
        PowerPointTable table = Assert.Single(Assert.Single(presentation.Slides).Tables);
        PowerPointTableCell header = table.GetCell(0, 0);
        PowerPointTableCell body = table.GetCell(1, 0);

        Assert.Equal("ABCDEF", header.FillColor);
        Assert.All(header.Runs, run => Assert.True(run.Bold));
        Assert.Equal("123456", header.Runs[0].Color);
        Assert.Equal(DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center, header.HorizontalAlignment);
        Assert.True(Assert.Single(body.Runs, run => run.Text.Contains("Bold", StringComparison.Ordinal)).Bold);
        PowerPointTextRun linked = Assert.Single(body.Runs,
            run => run.Text.Contains("linked", StringComparison.Ordinal));
        Assert.True(linked.Italic);
        Assert.Equal(new Uri("https://example.com/item"), linked.Hyperlink);
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void ExcelHtml_GenericMixedDocumentPreservesNarrativeAndTablesOnSeparateSheets() {
        HtmlToExcelResult result = HtmlConversionDocument.Parse("""
            <section><h1>Quarterly</h1><p>Executive narrative</p>
            <table><caption>Metrics</caption><tr><th>Name</th><th>Value</th></tr><tr><td>Revenue</td><td>42</td></tr></table></section>
            """).ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(2, workbook.Sheets.Count);
        ExcelSheet narrative = Assert.Single(workbook.Sheets, sheet => sheet.Name == "Imported");
        ExcelSheet metrics = Assert.Single(workbook.Sheets, sheet => sheet.Name == "Metrics");
        Assert.Contains("Executive narrative", Enumerable.Range(1, 8).Select(row => narrative.CellAt(row, 1).GetValue<string>()));
        Assert.Equal("Revenue", metrics.CellAt(2, 1).GetValue<string>());
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void GenericNativeAdaptersRetainInlineParagraphImages() {
        string html = "<section><h1>Evidence</h1><p>Before <strong>image</strong> <img src='" + PixelPng
            + "' alt='Inline evidence' width='24' height='30'> after</p></section>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        HtmlToExcelResult excelResult = source.ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = excelResult.Value;
        Assert.Equal("Inline evidence", Assert.Single(Assert.Single(workbook.Sheets).Images).Description);

        HtmlToPowerPointResult powerPointResult = source.ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = powerPointResult.Value;
        Assert.Equal("Inline evidence", Assert.Single(Assert.Single(presentation.Slides).Pictures).AltText);

        HtmlToOneNoteSectionResult oneNoteResult = source.ToOneNoteSectionResult();
        OneNotePage page = Assert.Single(oneNoteResult.Value.Pages);
        Assert.Contains(Assert.Single(page.Outlines).Children.OfType<OneNoteImage>(), image => image.AltText == "Inline evidence");
        Assert.Equal(1, excelResult.Images);
        Assert.Equal(1, powerPointResult.Pictures);
        Assert.Equal(1, oneNoteResult.Images);
    }

    [Fact]
    public void ExcelHtml_GenericRichRunsUseNormalizedBoundedText() {
        string html = "<p>A<strong>" + new string(' ', 40_000) + "B</strong></p>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;
        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        ExcelCell cell = Enumerable.Range(1, 4).Select(row => sheet.CellAt(row, 1))
            .First(candidate => candidate.GetValue<string>() == "A B");

        Assert.Equal("A B", string.Concat(cell.GetRichText().Select(run => run.Text)));
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
    }
}
