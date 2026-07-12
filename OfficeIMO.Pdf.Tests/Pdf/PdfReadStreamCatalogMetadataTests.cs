using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
    [Fact]
    public void RewriteApis_PreserveCatalogViewSettings() {
        byte[] catalogViewSettingPdf = BuildCatalogViewSettingPdf();
        byte[] twoPageCatalogViewSettingPdf = BuildTwoPageCatalogViewSettingPdf();

        AssertCatalogViewSettings(PdfPageExtractor.ExtractPages(catalogViewSettingPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(catalogViewSettingPdf);
        Assert.Single(splitPages);
        AssertCatalogViewSettings(splitPages[0]);
        AssertCatalogViewSettings(PdfPageEditor.DeletePages(twoPageCatalogViewSettingPdf, 2));
        AssertCatalogViewSettings(PdfPageEditor.ReorderPages(twoPageCatalogViewSettingPdf, 2, 1));
        AssertCatalogViewSettings(PdfPageEditor.RotatePages(catalogViewSettingPdf, 90));
        AssertCatalogViewSettings(PdfMetadataEditor.UpdateMetadata(catalogViewSettingPdf, title: "Updated"));
        AssertCatalogViewSettings(PdfMerger.Merge(catalogViewSettingPdf));
        AssertCatalogViewSettings(PdfStamper.StampText(catalogViewSettingPdf, "STAMP"));

        static void AssertCatalogViewSettings(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/PageMode /FullScreen", text, StringComparison.Ordinal);
            Assert.Contains("/PageLayout /TwoColumnLeft", text, StringComparison.Ordinal);

            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.True(info.HasCatalogViewSettings);
            Assert.Equal("FullScreen", info.CatalogPageMode);
            Assert.Equal("TwoColumnLeft", info.CatalogPageLayout);
        }
    }

    [Fact]
    public void RewriteApis_PreserveCatalogLanguageAndVersion() {
        byte[] catalogIdentityPdf = BuildCatalogIdentityPdf();
        byte[] twoPageCatalogIdentityPdf = BuildTwoPageCatalogIdentityPdf();

        AssertCatalogIdentity(PdfPageExtractor.ExtractPages(catalogIdentityPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(catalogIdentityPdf);
        Assert.Single(splitPages);
        AssertCatalogIdentity(splitPages[0]);
        AssertCatalogIdentity(PdfPageEditor.DeletePages(twoPageCatalogIdentityPdf, 2));
        AssertCatalogIdentity(PdfPageEditor.ReorderPages(twoPageCatalogIdentityPdf, 2, 1));
        AssertCatalogIdentity(PdfPageEditor.RotatePages(catalogIdentityPdf, 90));
        AssertCatalogIdentity(PdfMetadataEditor.UpdateMetadata(catalogIdentityPdf, title: "Updated"));
        AssertCatalogIdentity(PdfMerger.Merge(catalogIdentityPdf));
        AssertCatalogIdentity(PdfStamper.StampText(catalogIdentityPdf, "STAMP"));

        static void AssertCatalogIdentity(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Version /1.7", text, StringComparison.Ordinal);
            Assert.Contains("/Lang (pl-PL)", text, StringComparison.Ordinal);

            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.Equal("1.7", info.CatalogVersion);
            Assert.Equal("pl-PL", info.CatalogLanguage);
        }
    }

    [Fact]
    public void RewriteApis_PreserveCatalogUri() {
        byte[] catalogUriPdf = BuildCatalogUriPdf();
        byte[] twoPageCatalogUriPdf = BuildTwoPageCatalogUriPdf();

        AssertCatalogUri(PdfPageExtractor.ExtractPages(catalogUriPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(catalogUriPdf);
        Assert.Single(splitPages);
        AssertCatalogUri(splitPages[0]);
        AssertCatalogUri(PdfPageEditor.DeletePages(twoPageCatalogUriPdf, 2));
        AssertCatalogUri(PdfPageEditor.ReorderPages(twoPageCatalogUriPdf, 2, 1));
        AssertCatalogUri(PdfPageEditor.RotatePages(catalogUriPdf, 90));
        AssertCatalogUri(PdfMetadataEditor.UpdateMetadata(catalogUriPdf, title: "Updated"));
        AssertCatalogUri(PdfMerger.Merge(catalogUriPdf));
        AssertCatalogUri(PdfStamper.StampText(catalogUriPdf, "STAMP"));

        static void AssertCatalogUri(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/URI << /Base (https://evotec.xyz/docs/) >>", text, StringComparison.Ordinal);
            Assert.True(PdfInspector.Probe(output).HasCatalogUri);
        }
    }

    [Fact]
    public void RewriteApis_RejectComplexCatalogUriWithClearUnsupportedDiagnostic() {
        byte[] catalogUriPdf = BuildComplexCatalogUriPdf();

        static void AssertCatalogUri(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertCatalogUri(() => PdfPageExtractor.ExtractPages(catalogUriPdf, 1));
        AssertCatalogUri(() => PdfPageExtractor.SplitPages(catalogUriPdf));
        AssertCatalogUri(() => PdfPageEditor.DeletePages(catalogUriPdf, 1));
        AssertCatalogUri(() => PdfMetadataEditor.UpdateMetadata(catalogUriPdf, title: "Updated"));
        AssertCatalogUri(() => PdfMerger.Merge(catalogUriPdf));
        AssertCatalogUri(() => PdfStamper.StampText(catalogUriPdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_PreservePageLabels() {
        byte[] pageLabelPdf = BuildPageLabelPdf();
        byte[] twoPagePageLabelPdf = BuildTwoPageLabelPdf();

        AssertPageLabels(PdfPageExtractor.ExtractPages(pageLabelPdf, 1), "/Nums [ 0 << /S /D /St 1 >> ]");
        AssertPageLabels(PdfPageExtractor.ExtractPages(twoPagePageLabelPdf, 2), "/Nums [ 0 << /S /D /St 2 >> ]");
        var splitPages = PdfPageExtractor.SplitPages(pageLabelPdf);
        Assert.Single(splitPages);
        AssertPageLabels(splitPages[0], "/Nums [ 0 << /S /D /St 1 >> ]");
        var twoPageSplitPages = PdfPageExtractor.SplitPages(twoPagePageLabelPdf);
        Assert.Equal(2, twoPageSplitPages.Count);
        AssertPageLabels(twoPageSplitPages[0], "/Nums [ 0 << /S /D /St 1 >> ]");
        AssertPageLabels(twoPageSplitPages[1], "/Nums [ 0 << /S /D /St 2 >> ]");
        AssertPageLabels(PdfPageEditor.DeletePages(twoPagePageLabelPdf, 2), "/Nums [ 0 << /S /D /St 1 >> ]");
        AssertPageLabels(PdfPageEditor.DeletePages(twoPagePageLabelPdf, 1), "/Nums [ 0 << /S /D /St 2 >> ]");
        AssertPageLabels(PdfPageEditor.ReorderPages(twoPagePageLabelPdf, 2, 1), "/Nums [ 0 << /S /D /St 2 >> 1 << /S /D /St 1 >> ]");
        AssertPageLabels(PdfPageEditor.RotatePages(pageLabelPdf, 90), "/Nums [ 0 << /S /D /St 1 >> ]");
        AssertPageLabels(PdfMetadataEditor.UpdateMetadata(pageLabelPdf, title: "Updated"), "/Nums [ 0 << /S /D /St 1 >> ]");
        AssertPageLabels(PdfMerger.Merge(pageLabelPdf), "/Nums [ 0 << /S /D /St 1 >> ]");
        AssertPageLabels(PdfStamper.StampText(pageLabelPdf, "STAMP"), "/Nums [ 0 << /S /D /St 1 >> ]");

        static void AssertPageLabels(byte[] output, string expectedNums) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/PageLabels ", text, StringComparison.Ordinal);
            Assert.Contains(expectedNums, text, StringComparison.Ordinal);

            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.True(info.HasPageLabels);
            Assert.True(info.HasReadablePageLabels);
            Assert.NotEmpty(info.PageLabels);
            Assert.All(info.PageLabels, label => Assert.Equal("D", label.Style));
            Assert.All(info.PageLabels, label => Assert.Null(label.Prefix));
            Assert.All(info.PageLabels, label => Assert.True(label.StartNumber >= 1));
        }
    }

    [Fact]
    public void RewriteApis_ReindexPageLabelsUsingTrailerRootPageTreeWhenStaleCatalogsExist() {
        byte[] pdf = BuildStaleCatalogPageLabelsPdf();

        byte[] output = PdfPageExtractor.ExtractPages(pdf, 2);

        string text = System.Text.Encoding.ASCII.GetString(output);
        Assert.Contains("/PageLabels ", text, StringComparison.Ordinal);
        Assert.Contains("/Nums [ 0 << /S /D /St 10 >> ]", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/S /r", text, StringComparison.Ordinal);
        PdfPageLabel label = Assert.Single(PdfInspector.Inspect(output).PageLabels);
        Assert.Equal(0, label.StartPageIndex);
        Assert.Equal("D", label.Style);
        Assert.Equal(10, label.StartNumber);
    }

    [Fact]
    public void RewriteApis_RejectComplexPageLabelsWithClearUnsupportedDiagnostic() {
        byte[] pageLabelPdf = BuildComplexPageLabelPdf();

        static void AssertPageLabels(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("page labels", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        AssertPageLabels(() => PdfPageExtractor.ExtractPages(pageLabelPdf, 1));
        AssertPageLabels(() => PdfPageExtractor.SplitPages(pageLabelPdf));
        AssertPageLabels(() => PdfPageEditor.DeletePages(pageLabelPdf, 1));
        AssertPageLabels(() => PdfMetadataEditor.UpdateMetadata(pageLabelPdf, title: "Updated"));
        AssertPageLabels(() => PdfMerger.Merge(pageLabelPdf));
        AssertPageLabels(() => PdfStamper.StampText(pageLabelPdf, "STAMP"));
    }


}
