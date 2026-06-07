using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
    [Fact]
    public void RewriteApis_PreserveSimpleOutlinePdfsForCopiedPages() {
        byte[] outline = BuildOutlinePdf();

        AssertOutline(PdfPageExtractor.ExtractPages(outline, 1));
        var splitPages = PdfPageExtractor.SplitPages(outline);
        Assert.Single(splitPages);
        AssertOutline(splitPages[0]);
        AssertOutline(PdfPageEditor.ReorderPages(outline, 1));
        AssertOutline(PdfPageEditor.RotatePages(outline, 90));
        AssertOutline(PdfMetadataEditor.UpdateMetadata(outline, title: "Updated"));
        AssertOutline(PdfMerger.Merge(outline));
        AssertOutline(PdfStamper.StampText(outline, "STAMP"));

        static void AssertOutline(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Outlines ", text, StringComparison.Ordinal);
            Assert.Contains("/Type /Outlines", text, StringComparison.Ordinal);
            Assert.Contains("/Title (Executive summary)", text, StringComparison.Ordinal);

            PdfDocumentInfo info = PdfInspector.Inspect(output);
            PdfOutlineItem item = Assert.Single(info.Outlines);
            Assert.Equal("Executive summary", item.Title);
            Assert.Equal(1, item.PageNumber);
        }
    }

    [Fact]
    public void RewriteApis_DropSimpleOutlinesWhenDestinationPagesAreNotCopied() {
        byte[] outline = BuildTwoPageOutlinePdf();

        byte[] output = PdfPageEditor.DeletePages(outline, 2);

        string text = System.Text.Encoding.ASCII.GetString(output);
        Assert.DoesNotContain("/Outlines ", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/PageMode /UseOutlines", text, StringComparison.Ordinal);
        Assert.Empty(PdfInspector.Inspect(output).Outlines);
    }

    [Fact]
    public void ReadApis_ResolveOutlineIndirectDestinations() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildIndirectOutlineDestinationPdf());

        PdfOutlineItem item = Assert.Single(info.Outlines);
        Assert.Equal("Indirect destination", item.Title);
        Assert.Equal(1, item.PageNumber);
        Assert.Equal(144d, item.DestinationTop);
    }

    [Fact]
    public void ReadApis_ResolveOutlineGoToActionDestinations() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildGoToActionOutlinePdf());

        PdfOutlineItem item = Assert.Single(info.Outlines);
        Assert.Equal("Chapter 1", item.Title);
        Assert.Equal(1, item.PageNumber);
        Assert.Equal(200d, item.DestinationTop);
        Assert.Equal(PdfOpenActionDestinationMode.Xyz, item.DestinationMode);
    }

    [Fact]
    public void ReadApis_ResolveOutlineFitHorizontalDestinations() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildFitHorizontalOutlinePdf());

        PdfOutlineItem item = Assert.Single(info.Outlines);
        Assert.Equal("Fit horizontal", item.Title);
        Assert.Equal(1, item.PageNumber);
        Assert.Equal(144d, item.DestinationTop);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, item.DestinationMode);
    }

    [Fact]
    public void ReadApis_ResolveOutlineNamedDestinationTargets() {
        AssertOutline(BuildDirectNamedDestinationOutlinePdf(), "Direct named destination", 200d);
        AssertOutline(BuildNameTreeNamedDestinationOutlinePdf(), "Name-tree named destination", 188d);
        AssertOutline(BuildGoToActionNamedDestinationOutlinePdf(), "Action named destination", 176d);

        static void AssertOutline(byte[] pdf, string expectedTitle, double expectedTop) {
            PdfDocumentInfo info = PdfInspector.Inspect(pdf);

            PdfOutlineItem item = Assert.Single(info.Outlines);
            Assert.Equal(expectedTitle, item.Title);
            Assert.Equal(1, item.PageNumber);
            Assert.Equal(expectedTop, item.DestinationTop);
            Assert.Equal(PdfOpenActionDestinationMode.Xyz, item.DestinationMode);
            PdfNamedDestination destination = Assert.Single(info.NamedDestinations);
            Assert.Equal("Chapter1", destination.Name);
            Assert.Equal(1, destination.PageNumber);
            Assert.Equal(expectedTop, destination.DestinationTop);
            Assert.Equal(PdfOpenActionDestinationMode.Xyz, destination.DestinationMode);
        }
    }

    [Fact]
    public void ReadApis_ResolveOutlineNamedDestinationTargetsByTokenNamespace() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildMixedNamedDestinationOutlinePdf());

        Assert.Equal(3, info.Outlines.Count);
        Assert.Equal("Direct name destination", info.Outlines[0].Title);
        Assert.Equal(144d, info.Outlines[0].DestinationTop);
        Assert.Equal("Name-tree string destination", info.Outlines[1].Title);
        Assert.Equal(188d, info.Outlines[1].DestinationTop);
        Assert.Equal("Dictionary string destination", info.Outlines[2].Title);
        Assert.Equal(188d, info.Outlines[2].DestinationTop);
    }

    [Fact]
    public void RewriteApis_PreserveGoToActionOutlinePdfsForCopiedPages() {
        byte[] outline = BuildGoToActionOutlinePdf();

        AssertOutline(PdfPageExtractor.ExtractPages(outline, 1));
        var splitPages = PdfPageExtractor.SplitPages(outline);
        Assert.Single(splitPages);
        AssertOutline(splitPages[0]);
        AssertOutline(PdfPageEditor.ReorderPages(outline, 1));
        AssertOutline(PdfPageEditor.RotatePages(outline, 90));
        AssertOutline(PdfMetadataEditor.UpdateMetadata(outline, title: "Updated"));
        AssertOutline(PdfMerger.Merge(outline));
        AssertOutline(PdfStamper.StampText(outline, "STAMP"));

        static void AssertOutline(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Outlines ", text, StringComparison.Ordinal);
            Assert.Contains("/A << /S /GoTo /D [", text, StringComparison.Ordinal);
            Assert.Contains("/Title (Chapter 1)", text, StringComparison.Ordinal);

            PdfOutlineItem item = Assert.Single(PdfInspector.Inspect(output).Outlines);
            Assert.Equal("Chapter 1", item.Title);
            Assert.Equal(1, item.PageNumber);
            Assert.Equal(200d, item.DestinationTop);
        }
    }

    [Fact]
    public void RewriteApis_PreserveGoToActionOutlinePdfsWithIndirectDestinationsForCopiedPages() {
        byte[] outline = BuildGoToActionIndirectDestinationPdf();

        PdfDocumentPreflight preflight = PdfInspector.Preflight(outline);
        Assert.True(preflight.CanRewrite);
        Assert.False(preflight.HasRewriteBlocker(PdfRewriteBlockerKind.Outlines));

        byte[] output = PdfPageExtractor.ExtractPages(outline, 1);

        PdfOutlineItem item = Assert.Single(PdfInspector.Inspect(output).Outlines);
        Assert.Equal("Indirect GoTo action", item.Title);
        Assert.Equal(1, item.PageNumber);
        Assert.Equal(188d, item.DestinationTop);
    }

    [Fact]
    public void RewriteApis_PreserveGoToActionOutlinePdfsWithDictionaryDestinationsForCopiedPages() {
        byte[] outline = BuildGoToActionDictionaryDestinationPdf();

        PdfDocumentPreflight preflight = PdfInspector.Preflight(outline);
        Assert.True(preflight.CanRewrite);
        Assert.False(preflight.HasRewriteBlocker(PdfRewriteBlockerKind.Outlines));

        byte[] output = PdfPageExtractor.ExtractPages(outline, 1);

        PdfOutlineItem item = Assert.Single(PdfInspector.Inspect(output).Outlines);
        Assert.Equal("Dictionary GoTo action", item.Title);
        Assert.Equal(1, item.PageNumber);
        Assert.Equal(188d, item.DestinationTop);
    }

    [Fact]
    public void Preflight_BlocksCyclicGoToActionOutlineDestinations() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCyclicGoToActionDestinationPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.Contains(report.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.Outlines);
    }

    [Fact]
    public void RewriteApis_RejectComplexOutlinePdfsWithClearUnsupportedDiagnostic() {
        byte[] outline = BuildUriActionOutlinePdf();

        static void AssertOutline(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertOutline(() => PdfPageExtractor.ExtractPages(outline, 1));
        AssertOutline(() => PdfPageExtractor.SplitPages(outline));
        AssertOutline(() => PdfPageEditor.DeletePages(outline, 1));
        AssertOutline(() => PdfMetadataEditor.UpdateMetadata(outline, title: "Updated"));
        AssertOutline(() => PdfMerger.Merge(outline));
        AssertOutline(() => PdfStamper.StampText(outline, "STAMP"));
    }


}
