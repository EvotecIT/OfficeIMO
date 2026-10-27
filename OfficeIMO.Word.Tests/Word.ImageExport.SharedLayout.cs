using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordSharedPageLayoutTests {
    [Theory]
    [InlineData("word-business-delivery-summary.docx", OfficeImageExportFormat.Png)]
    [InlineData("word-business-delivery-summary.docx", OfficeImageExportFormat.Jpeg)]
    [InlineData("word-business-delivery-summary.docx", OfficeImageExportFormat.Svg)]
    [InlineData("native-word-report.docx", OfficeImageExportFormat.Png)]
    [InlineData("native-word-report.docx", OfficeImageExportFormat.Jpeg)]
    [InlineData("native-word-report.docx", OfficeImageExportFormat.Svg)]
    public void NativeFixturesUseExactlyThePdfPageScenes(string fixture, OfficeImageExportFormat format) {
        using WordDocument document = WordDocument.Load(Path.Combine(AppContext.BaseDirectory, "WordLayoutFixtures", fixture));
        var expected = document.ConvertToPdfWithEvidence().ExportImages(format, new PdfImageExportOptions { Scale = .25D });
        var actual = document.ExportImages(format, new WordImageExportOptions { Scale = .25D });
        Assert.Equal(expected.Count, actual.Count);
        for (int index = 0; index < actual.Count; index++) Assert.Equal(expected[index].Bytes, actual[index].Bytes);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Svg)]
    public void WordImagesMatchPdfPagesIncludingStylesTablesAndOverflow(OfficeImageExportFormat format) {
        using WordDocument document = CreateDocument();
        var imageOptions = new WordImageExportOptions { Scale = .5D };
        var expected = document.ConvertToPdfWithEvidence().ExportImages(format, new PdfImageExportOptions { Scale = .5D });
        var actual = document.ExportImages(format, imageOptions);
        Assert.True(actual.Count > 1);
        Assert.Equal(expected.Count, actual.Count);
        for (int index = 0; index < actual.Count; index++) {
            Assert.Equal(expected[index].Width, actual[index].Width);
            Assert.Equal(expected[index].Height, actual[index].Height);
            Assert.Equal(expected[index].Bytes, actual[index].Bytes);
        }
        using var pdf = new MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var read = PdfReadDocument.Open(pdf);
        var snapshots = document.CreateVisualSnapshots();
        Assert.Equal(read.Pages.Count, document.GetEstimatedImagePageCount());
        for (int index = 0; index < snapshots.Count; index++) {
            Assert.Equal(read.Pages[index].GetTextSpans(includeArtifactText: true).Select(Geometry), snapshots[index].TextSpans.Select(Geometry));
            Assert.DoesNotContain(snapshots[index].TextSpans, span => span.Text.Contains('…'));
        }
        string text = string.Join("", snapshots.SelectMany(page => page.TextSpans).Select(span => span.Text));
        Assert.Contains("END_OF_LONG_PARAGRAPH", text);
        Assert.Contains("END_OF_TABLE_CELL", text);
        Assert.Contains(snapshots[0].TextSpans, span => span.FontSize > 18D && span.IsBold);
        if (format == OfficeImageExportFormat.Svg)
            Assert.Contains("#FF0000", System.Text.Encoding.UTF8.GetString(actual.SelectMany(page => page.Bytes).ToArray()), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WordImageSelectionStreamingAndWrappersUseTheSamePages() {
        using WordDocument document = CreateDocument();
        var options = new WordImageExportOptions { PageIndex = 1, PageCount = 1, Scale = .3D };
        var page = Assert.Single(document.ExportImages(OfficeImageExportFormat.Png, options));
        Assert.Equal(page.Bytes, document.ToPng(options));
        var single = document.ExportImage(OfficeImageExportFormat.Png, options);
        AssertWordIdentity(single, 2, 0, 1);
        AssertWordIdentity(page, 2, 0, 1);
        var delivered = new System.Collections.Generic.List<OfficeImageExportResult>();
        document.ExportImages(OfficeImageExportFormat.Png, delivered.Add, options);
        Assert.Equal(page.Bytes, Assert.Single(delivered).Bytes);
        AssertWordIdentity(delivered[0], 2, 0, 1);
        var selected = document.ExportImages(OfficeImageExportFormat.Svg, new WordImageExportOptions { PageIndex = 1, PageCount = 2 });
        Assert.Equal(2, selected.Count);
        for (int index = 0; index < selected.Count; index++) AssertWordIdentity(selected[index], index + 2, index, 2);
        Assert.Equal(1, Assert.Single(document.CreateVisualSnapshots(options)).PageIndex);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.ToPng(new WordImageExportOptions { PageIndex = 999 }));
    }

    [Fact]
    public void SharedLayoutHonorsConversionSettingsAndClonesThem() {
        using WordDocument document = CreateDocument();
        var options = new WordImageExportOptions {
            LayoutOptions = new WordPdfExportOptions { PageSize = new OfficeIMO.Pdf.PdfPageSize(400, 500), ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() }
        };
        WordImageExportOptions clone = options.Clone();
        options.LayoutOptions.PageSize = new OfficeIMO.Pdf.PdfPageSize(700, 800);
        var expected = document.ConvertToPdfWithEvidence(clone.LayoutOptions).ExportImages(OfficeImageExportFormat.Svg);
        var actual = document.ExportImages(OfficeImageExportFormat.Svg, clone);
        Assert.Equal(expected.Select(page => page.Bytes), actual.Select(page => page.Bytes), ByteArrayComparer.Instance);
        Assert.Equal(400D, document.CreateVisualSnapshot(clone).Width);
    }

    [Fact]
    public void SharedConversionReportsOmissionsAndStrictImagesRejectThem() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Visible text before the unsupported chart.");
        var chart = document.AddChart(OfficeChartType.Pie, new OfficeChartData(new[] { "A" },
            new[] { new OfficeSeries("Values", new[] { 1D }) }));
        chart.ChartPart!.ChartSpace!.RemoveAllChildren();
        var conversion = document.ConvertToPdfWithEvidence();
        Assert.Contains(conversion.Warnings, warning => warning.Code == "native-body-chart-unsupported"
            && warning.LossKind == OfficeConversionLossKind.Omission);
        var options = new WordImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoOmissions = true } };
        Assert.Throws<OfficeImageExportPolicyException>(() => document.ExportImage(OfficeImageExportFormat.Png, options));
        Assert.Throws<OfficeImageExportPolicyException>(() => conversion.ExportImages(OfficeImageExportFormat.Png,
            new PdfImageExportOptions { Policy = options.Policy }));
    }

    [Fact]
    public void SharedLayoutRetainsCancellationTimeoutAndBatchBudgets() {
        using WordDocument document = CreateDocument();
        Assert.Throws<OperationCanceledException>(() => document.ExportImage(OfficeImageExportFormat.Png, cancellationToken: new CancellationToken(true)));
        Assert.Throws<OfficeImageExportTimeoutException>(() => document.ExportImage(OfficeImageExportFormat.Svg,
            new WordImageExportOptions { RenderTimeout = TimeSpan.FromTicks(1) }));
        Assert.Throws<OfficeImageExportBatchLimitException>(() => document.ExportImages(OfficeImageExportFormat.Png,
            new WordImageExportOptions { MaximumOutputCount = 1 }));
        Assert.Throws<OfficeImageExportBatchLimitException>(() => document.ExportImage(OfficeImageExportFormat.Png,
            new WordImageExportOptions { MaximumTotalEncodedBytes = 8 }));
        Assert.Throws<OperationCanceledException>(() => document.CreateVisualSnapshots(null, new CancellationToken(true)));
    }

    private static string Geometry(PdfTextSpan span) =>
        $"{span.Text}|{span.X:R}|{span.Y:R}|{span.Advance:R}|{span.FontSize:R}|{span.BaseFont}";

    private static void AssertWordIdentity(OfficeImageExportResult result, int pageNumber, int sequenceIndex, int sequenceCount) {
        Assert.Equal("Page " + pageNumber, result.Name);
        Assert.Equal("Word document", result.Source);
        Assert.Equal(sequenceIndex, result.SequenceIndex);
        Assert.Equal(sequenceCount, result.SequenceCount);
    }

    private static WordDocument CreateDocument() {
        WordDocument document = WordDocument.Create();
        document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Append(new Style(
            new StyleName { Val = "Shared heading" },
            new StyleRunProperties(new FontSize { Val = "41" }, new Bold())) {
                Type = StyleValues.Paragraph, StyleId = "SharedHeading"
            });
        document.AddParagraph("Shared heading").SetStyleId("SharedHeading");
        document.AddParagraph(string.Join(" ", Enumerable.Repeat("Full paragraph wraps onto the following page.", 160)) + " END_OF_LONG_PARAGRAPH");
        WordTable table = document.AddTable(2, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = string.Join(" ", Enumerable.Repeat("A table cell wraps completely.", 100)) + " END_OF_TABLE_CELL";
        table._tableProperties!.TableBorders = new TableBorders(new TopBorder { Val = BorderValues.Single, Color = "FF0000", Size = 16 });
        document.AddPageBreak();
        document.AddParagraph("Final page");
        return document;
    }

    private sealed class ByteArrayComparer : System.Collections.Generic.IEqualityComparer<byte[]> {
        internal static ByteArrayComparer Instance { get; } = new();
        public bool Equals(byte[]? left, byte[]? right) => left != null && right != null && left.SequenceEqual(right);
        public int GetHashCode(byte[] value) => value.Length;
    }
}
