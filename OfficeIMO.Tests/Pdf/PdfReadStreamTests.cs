using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfReadStreamTests {
    [Fact]
    public void PdfReadDocument_Load_ReadsFromCurrentStreamPosition() {
        byte[] pdf = BuildPdf();
        using var stream = BuildPrefixedStream(pdf);
        stream.Position = 5;

        PdfReadDocument document = PdfReadDocument.Load(stream);

        Assert.Single(document.Pages);
        Assert.Equal("Stream read", document.Metadata.Title);
        Assert.Contains("Streamreadabletext", document.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsFromCurrentStreamPosition() {
        byte[] pdf = BuildPdf();
        using var stream = BuildPrefixedStream(pdf);
        stream.Position = 5;

        string text = PdfTextExtractor.ExtractAllText(stream);

        Assert.Contains("Stream readable text", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_GetMetadata_ReadsFromPathAndStream() {
        byte[] pdf = BuildPdf();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-read-stream-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(path, pdf);
            using var stream = BuildPrefixedStream(pdf);
            stream.Position = 5;

            var fromPath = PdfTextExtractor.GetMetadata(path);
            var fromStream = PdfTextExtractor.GetMetadata(stream);

            Assert.Equal("Stream read", fromPath.Title);
            Assert.Equal("OfficeIMO", fromPath.Author);
            Assert.Equal(fromPath.Title, fromStream.Title);
            Assert.Equal(fromPath.Author, fromStream.Author);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ReadStreamApis_RejectNullAndUnreadableStreams() {
        Assert.Throws<ArgumentNullException>(() => PdfReadDocument.Load((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.GetMetadata((Stream)null!));

        using var unreadable = new WriteOnlyStream();

        Assert.Throws<ArgumentException>(() => PdfReadDocument.Load(unreadable));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText(unreadable));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.GetMetadata(unreadable));
    }

    [Fact]
    public void ReadPathApis_RejectNullAndWhitespacePaths() {
        Assert.Throws<ArgumentNullException>(() => PdfReadDocument.Load((string)null!));
        Assert.Throws<ArgumentException>(() => PdfReadDocument.Load(" "));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText((string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText(" "));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPage((string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractTextByPage(" "));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.GetMetadata((string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.GetMetadata(" "));

        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImages((string)null!));
        Assert.Throws<ArgumentException>(() => PdfImageExtractor.ExtractImages(" "));
    }

    [Fact]
    public void ReadApis_RejectEncryptedPdfsWithClearUnsupportedDiagnostic() {
        byte[] encrypted = BuildEncryptedPdfMarker();

        AssertEncrypted(() => PdfReadDocument.Load(encrypted));
        AssertEncrypted(() => PdfTextExtractor.ExtractAllText(encrypted));
        AssertEncrypted(() => PdfTextExtractor.GetMetadata(encrypted));
        AssertEncrypted(() => PdfImageExtractor.ExtractImages(encrypted));
        AssertEncrypted(() => PdfPageExtractor.ExtractPages(encrypted, 1));

        static void AssertEncrypted(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("Encrypted PDF files are not supported by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_RejectSignedPdfsWithClearUnsupportedDiagnostic() {
        byte[] signed = BuildSignedPdfMarker();

        AssertSigned(() => PdfPageExtractor.ExtractPages(signed, 1));
        AssertSigned(() => PdfPageExtractor.SplitPages(signed));
        AssertSigned(() => PdfPageEditor.DeletePages(signed, 1));
        AssertSigned(() => PdfMetadataEditor.UpdateMetadata(signed, title: "Updated"));
        AssertSigned(() => PdfMerger.Merge(signed));
        AssertSigned(() => PdfStamper.StampText(signed, "STAMP"));

        static void AssertSigned(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_RejectFormPdfsWithClearUnsupportedDiagnostic() {
        byte[] form = BuildFormPdfMarker();

        AssertForm(() => PdfPageExtractor.ExtractPages(form, 1));
        AssertForm(() => PdfPageExtractor.SplitPages(form));
        AssertForm(() => PdfPageEditor.DeletePages(form, 1));
        AssertForm(() => PdfMetadataEditor.UpdateMetadata(form, title: "Updated"));
        AssertForm(() => PdfMerger.Merge(form));
        AssertForm(() => PdfStamper.StampText(form, "STAMP"));

        static void AssertForm(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }
    }

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

    [Fact]
    public void RewriteApis_UseTrailerRootCatalogWhenStaleCatalogRevisionsExist() {
        byte[] pdf = BuildStaleCatalogRevisionPdf();

        PdfDocumentInfo inputInfo = PdfInspector.Inspect(pdf);
        Assert.Equal("SinglePage", inputInfo.CatalogPageLayout);
        Assert.False(inputInfo.HasReadablePageLabels);
        Assert.False(PdfInspector.Preflight(pdf).HasRewriteBlocker(PdfRewriteBlockerKind.PageLabels));

        byte[] output = PdfPageExtractor.ExtractPages(pdf, 1);

        string text = System.Text.Encoding.ASCII.GetString(output);
        Assert.Contains("/PageLayout /SinglePage", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/PageLayout /TwoColumnLeft", text, StringComparison.Ordinal);
        Assert.False(PdfInspector.Inspect(output).HasReadablePageLabels);
    }

    [Fact]
    public void ReadApis_UseTrailerRootCatalogForPagesAndOutlinesWhenStaleCatalogsExist() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildStaleCatalogWithDifferentPagesAndOutlinesPdf());

        PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(200d, page.Width);
        Assert.Equal(200d, page.Height);
        PdfOutlineItem outline = Assert.Single(info.Outlines);
        Assert.Equal("Current", outline.Title);
        Assert.Equal(1, outline.PageNumber);
        Assert.Equal("SinglePage", info.CatalogPageLayout);
    }

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
            var exception = Assert.Throws<NotSupportedException>(action);
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
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("page labels", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        AssertPageLabels(() => PdfPageExtractor.ExtractPages(pageLabelPdf, 1));
        AssertPageLabels(() => PdfPageExtractor.SplitPages(pageLabelPdf));
        AssertPageLabels(() => PdfPageEditor.DeletePages(pageLabelPdf, 1));
        AssertPageLabels(() => PdfMetadataEditor.UpdateMetadata(pageLabelPdf, title: "Updated"));
        AssertPageLabels(() => PdfMerger.Merge(pageLabelPdf));
        AssertPageLabels(() => PdfStamper.StampText(pageLabelPdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_PreserveDirectNamedDestinationsForCopiedPages() {
        byte[] namedDestinationPdf = BuildNamedDestinationPdf();
        byte[] twoPageNamedDestinationPdf = BuildTwoPageNamedDestinationPdf();

        AssertNamedDestinations(PdfPageExtractor.ExtractPages(namedDestinationPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(namedDestinationPdf);
        Assert.Single(splitPages);
        AssertNamedDestinations(splitPages[0]);
        AssertNamedDestinations(PdfPageEditor.DeletePages(twoPageNamedDestinationPdf, 2), containsSecondDestination: false);
        AssertNamedDestinations(PdfPageEditor.ReorderPages(twoPageNamedDestinationPdf, 2, 1), containsSecondDestination: true);
        AssertNamedDestinations(PdfPageEditor.RotatePages(namedDestinationPdf, 90));
        AssertNamedDestinations(PdfMetadataEditor.UpdateMetadata(namedDestinationPdf, title: "Updated"));
        AssertNamedDestinations(PdfMerger.Merge(namedDestinationPdf));
        AssertNamedDestinations(PdfStamper.StampText(namedDestinationPdf, "STAMP"));

        static void AssertNamedDestinations(byte[] output, bool containsSecondDestination = false) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Dests ", text, StringComparison.Ordinal);
            Assert.Contains("/Chapter1 [", text, StringComparison.Ordinal);
            Assert.Contains("/XYZ 0 200 0", text, StringComparison.Ordinal);
            if (containsSecondDestination) {
                Assert.Contains("/Chapter2 [", text, StringComparison.Ordinal);
                Assert.Contains("/Fit", text, StringComparison.Ordinal);
            } else {
                Assert.DoesNotContain("/Chapter2 [", text, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void RewriteApis_PreserveNamedDestinationNameTreesForCopiedPages() {
        byte[] namedDestinationPdf = BuildNamedDestinationNameTreePdf();
        byte[] twoPageNamedDestinationPdf = BuildTwoPageNamedDestinationNameTreePdf();

        AssertNamedDestinations(PdfPageExtractor.ExtractPages(namedDestinationPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(namedDestinationPdf);
        Assert.Single(splitPages);
        AssertNamedDestinations(splitPages[0]);
        AssertNamedDestinations(PdfPageEditor.DeletePages(twoPageNamedDestinationPdf, 2), containsSecondDestination: false);
        AssertNamedDestinations(PdfPageEditor.ReorderPages(twoPageNamedDestinationPdf, 2, 1), containsSecondDestination: true);
        AssertNamedDestinations(PdfPageEditor.RotatePages(namedDestinationPdf, 90));
        AssertNamedDestinations(PdfMetadataEditor.UpdateMetadata(namedDestinationPdf, title: "Updated"));
        AssertNamedDestinations(PdfMerger.Merge(namedDestinationPdf));
        AssertNamedDestinations(PdfStamper.StampText(namedDestinationPdf, "STAMP"));

        static void AssertNamedDestinations(byte[] output, bool containsSecondDestination = false) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Names << /Dests << /Names [", text, StringComparison.Ordinal);
            Assert.Contains("(Chapter1)", text, StringComparison.Ordinal);
            Assert.Contains("/XYZ 0 200 0", text, StringComparison.Ordinal);
            if (containsSecondDestination) {
                Assert.Contains("(Chapter2)", text, StringComparison.Ordinal);
                Assert.Contains("/Fit", text, StringComparison.Ordinal);
            } else {
                Assert.DoesNotContain("(Chapter2)", text, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void RewriteApis_PreserveNamedDestinationNameTreeKidsForCopiedPages() {
        byte[] namedDestinationPdf = BuildNamedDestinationNameTreeWithKidsPdf();
        byte[] twoPageNamedDestinationPdf = BuildTwoPageNamedDestinationNameTreeWithKidsPdf();

        AssertNamedDestinations(PdfPageExtractor.ExtractPages(namedDestinationPdf, 1));
        AssertNamedDestinations(PdfPageEditor.DeletePages(twoPageNamedDestinationPdf, 2), containsSecondDestination: false);
        AssertNamedDestinations(PdfPageEditor.ReorderPages(twoPageNamedDestinationPdf, 2, 1), containsSecondDestination: true, chapter1Page: 2, chapter2Page: 1);
        AssertNamedDestinations(PdfMetadataEditor.UpdateMetadata(namedDestinationPdf, title: "Updated"));

        static void AssertNamedDestinations(byte[] output, bool containsSecondDestination = false, int chapter1Page = 1, int chapter2Page = 2) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Names << /Dests << /Names [", text, StringComparison.Ordinal);
            Assert.DoesNotContain("/Dests << /Kids", text, StringComparison.Ordinal);

            PdfDocumentInfo info = PdfInspector.Inspect(output);
            AssertDestination(info, "Chapter1", chapter1Page, 200);
            if (containsSecondDestination) {
                AssertDestination(info, "Chapter2", chapter2Page, null);
            } else {
                Assert.DoesNotContain(info.NamedDestinations, destination => destination.Name == "Chapter2");
            }
        }

        static void AssertDestination(PdfDocumentInfo info, string name, int pageNumber, double? top) {
            PdfNamedDestination destination = Assert.Single(info.NamedDestinations, item => item.Name == name);
            Assert.Equal(pageNumber, destination.PageNumber);
            Assert.Equal(top, destination.DestinationTop);
        }
    }

    [Fact]
    public void RewriteApis_PreserveSimpleLinkAnnotationsWithContentsMetadata() {
        byte[] linkedPdf = BuildTwoPageLinkAnnotationPdf();

        AssertSinglePageLink(PdfPageExtractor.ExtractPages(linkedPdf, 1), "https://evotec.xyz/first", "First link metadata");
        var splitPages = PdfPageExtractor.SplitPages(linkedPdf);
        Assert.Equal(2, splitPages.Count);
        AssertSinglePageLink(splitPages[0], "https://evotec.xyz/first", "First link metadata");
        AssertSinglePageLink(splitPages[1], "https://evotec.xyz/second", "Second link metadata");

        AssertSinglePageLink(PdfPageEditor.DeletePages(linkedPdf, 2), "https://evotec.xyz/first", "First link metadata");
        AssertTwoPageLinks(
            PdfPageEditor.ReorderPages(linkedPdf, 2, 1),
            ("https://evotec.xyz/second", "Second link metadata"),
            ("https://evotec.xyz/first", "First link metadata"));
        AssertTwoPageLinks(
            PdfPageEditor.RotatePages(linkedPdf, 90),
            ("https://evotec.xyz/first", "First link metadata"),
            ("https://evotec.xyz/second", "Second link metadata"));
        AssertTwoPageLinks(
            PdfMetadataEditor.UpdateMetadata(linkedPdf, title: "Updated"),
            ("https://evotec.xyz/first", "First link metadata"),
            ("https://evotec.xyz/second", "Second link metadata"));
        AssertTwoPageLinks(
            PdfMerger.Merge(linkedPdf),
            ("https://evotec.xyz/first", "First link metadata"),
            ("https://evotec.xyz/second", "Second link metadata"));
        AssertTwoPageLinks(
            PdfStamper.StampText(linkedPdf, "STAMP"),
            ("https://evotec.xyz/first", "First link metadata"),
            ("https://evotec.xyz/second", "Second link metadata"));

        static void AssertSinglePageLink(byte[] output, string uri, string contents) {
            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.Single(info.Pages);
            AssertPageLink(info.Pages[0], uri, contents);
        }

        static void AssertTwoPageLinks(byte[] output, (string Uri, string Contents) first, (string Uri, string Contents) second) {
            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.Equal(2, info.PageCount);
            Assert.Equal(2, info.LinkAnnotationCount);
            Assert.Equal(2, info.LinkAnnotations.Count);
            Assert.Equal(2, info.LinkUriCount);
            Assert.Equal(new[] { first.Uri, second.Uri }, info.LinkUris);
            AssertPageLink(info.Pages[0], first.Uri, first.Contents);
            AssertPageLink(info.Pages[1], second.Uri, second.Contents);
            Assert.Equal(first.Uri, info.LinkAnnotations[0].Uri);
            Assert.Equal(1, info.LinkAnnotations[0].PageNumber);
            Assert.Equal(second.Uri, info.LinkAnnotations[1].Uri);
            Assert.Equal(2, info.LinkAnnotations[1].PageNumber);
        }

        static void AssertPageLink(PdfPageInfo page, string uri, string contents) {
            var link = Assert.Single(page.LinkAnnotations);
            Assert.Equal(page.PageNumber, link.PageNumber);
            Assert.Equal(uri, link.Uri);
            Assert.Equal(contents, link.Contents);
            Assert.True(link.Width > 0);
            Assert.True(link.Height > 0);
            Assert.InRange(link.X1, 0, page.Width);
            Assert.InRange(link.X2, 0, page.Width);
            Assert.InRange(link.Y1, 0, page.Height);
            Assert.InRange(link.Y2, 0, page.Height);
        }
    }

    [Fact]
    public void RewriteApis_RejectComplexNamedDestinationNameTreesWithClearUnsupportedDiagnostic() {
        byte[] namedDestinationPdf = BuildComplexNamedDestinationNameTreePdf();

        static void AssertNamedDestinations(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertNamedDestinations(() => PdfPageExtractor.ExtractPages(namedDestinationPdf, 1));
        AssertNamedDestinations(() => PdfPageExtractor.SplitPages(namedDestinationPdf));
        AssertNamedDestinations(() => PdfPageEditor.DeletePages(namedDestinationPdf, 1));
        AssertNamedDestinations(() => PdfMetadataEditor.UpdateMetadata(namedDestinationPdf, title: "Updated"));
        AssertNamedDestinations(() => PdfMerger.Merge(namedDestinationPdf));
        AssertNamedDestinations(() => PdfStamper.StampText(namedDestinationPdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_PreserveDestinationOpenActionsForCopiedPages() {
        byte[] openActionPdf = BuildOpenActionPdf();
        byte[] twoPageOpenActionPdf = BuildTwoPageOpenActionPdf();

        AssertOpenAction(PdfPageExtractor.ExtractPages(openActionPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(openActionPdf);
        Assert.Single(splitPages);
        AssertOpenAction(splitPages[0]);
        AssertNoOpenAction(PdfPageEditor.DeletePages(twoPageOpenActionPdf, 2));
        AssertOpenAction(PdfPageEditor.ReorderPages(twoPageOpenActionPdf, 2, 1));
        AssertOpenAction(PdfPageEditor.RotatePages(openActionPdf, 90));
        AssertOpenAction(PdfMetadataEditor.UpdateMetadata(openActionPdf, title: "Updated"));
        AssertOpenAction(PdfMerger.Merge(openActionPdf));
        AssertOpenAction(PdfStamper.StampText(openActionPdf, "STAMP"));

        static void AssertOpenAction(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/OpenAction [", text, StringComparison.Ordinal);
            Assert.Contains("/Fit", text, StringComparison.Ordinal);

            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.True(info.HasOpenActions);
            Assert.True(info.HasReadableOpenAction);
            Assert.NotNull(info.OpenAction);
            Assert.Equal("Destination", info.OpenAction!.ActionType);
            Assert.Equal(1, info.OpenAction.PageNumber);
            Assert.Null(info.OpenAction.DestinationTop);
        }

        static void AssertNoOpenAction(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.DoesNotContain("/OpenAction ", text, StringComparison.Ordinal);
            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.False(info.HasOpenActions);
            Assert.False(info.HasReadableOpenAction);
            Assert.Null(info.OpenAction);
        }
    }

    [Fact]
    public void RewriteApis_PreserveGoToOpenActionDictionariesForCopiedPages() {
        byte[] openActionPdf = BuildOpenActionDictionaryPdf();
        byte[] twoPageOpenActionPdf = BuildTwoPageOpenActionDictionaryPdf();

        AssertOpenAction(PdfPageExtractor.ExtractPages(openActionPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(openActionPdf);
        Assert.Single(splitPages);
        AssertOpenAction(splitPages[0]);
        AssertNoOpenAction(PdfPageEditor.DeletePages(twoPageOpenActionPdf, 2));
        AssertOpenAction(PdfPageEditor.ReorderPages(twoPageOpenActionPdf, 2, 1));
        AssertOpenAction(PdfPageEditor.RotatePages(openActionPdf, 90));
        AssertOpenAction(PdfMetadataEditor.UpdateMetadata(openActionPdf, title: "Updated"));
        AssertOpenAction(PdfMerger.Merge(openActionPdf));
        AssertOpenAction(PdfStamper.StampText(openActionPdf, "STAMP"));

        static void AssertOpenAction(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/OpenAction <<", text, StringComparison.Ordinal);
            Assert.Contains("/S /GoTo", text, StringComparison.Ordinal);
            Assert.Contains("/D [", text, StringComparison.Ordinal);
            Assert.Contains("/Fit", text, StringComparison.Ordinal);

            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.True(info.HasOpenActions);
            Assert.True(info.HasReadableOpenAction);
            Assert.NotNull(info.OpenAction);
            Assert.Equal("GoTo", info.OpenAction!.ActionType);
            Assert.Equal(1, info.OpenAction.PageNumber);
            Assert.Null(info.OpenAction.DestinationTop);
        }

        static void AssertNoOpenAction(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.DoesNotContain("/OpenAction ", text, StringComparison.Ordinal);
            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.False(info.HasOpenActions);
            Assert.False(info.HasReadableOpenAction);
            Assert.Null(info.OpenAction);
        }
    }

    [Fact]
    public void RewriteApis_RejectComplexOpenActionDictionariesWithClearUnsupportedDiagnostic() {
        byte[] openActionPdf = BuildComplexOpenActionDictionaryPdf();

        static void AssertOpenActions(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertOpenActions(() => PdfPageExtractor.ExtractPages(openActionPdf, 1));
        AssertOpenActions(() => PdfPageExtractor.SplitPages(openActionPdf));
        AssertOpenActions(() => PdfPageEditor.DeletePages(openActionPdf, 1));
        AssertOpenActions(() => PdfMetadataEditor.UpdateMetadata(openActionPdf, title: "Updated"));
        AssertOpenActions(() => PdfMerger.Merge(openActionPdf));
        AssertOpenActions(() => PdfStamper.StampText(openActionPdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_PreserveSimpleViewerPreferences() {
        byte[] viewerPreferencePdf = BuildViewerPreferencePdf();

        AssertViewerPreferences(PdfPageExtractor.ExtractPages(viewerPreferencePdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(viewerPreferencePdf);
        Assert.Single(splitPages);
        AssertViewerPreferences(splitPages[0]);
        AssertViewerPreferences(PdfPageEditor.DeletePages(BuildTwoPageViewerPreferencePdf(), 2));
        AssertViewerPreferences(PdfPageEditor.ReorderPages(BuildTwoPageViewerPreferencePdf(), 2, 1));
        AssertViewerPreferences(PdfPageEditor.RotatePages(viewerPreferencePdf, 90));
        AssertViewerPreferences(PdfMetadataEditor.UpdateMetadata(viewerPreferencePdf, title: "Updated"));
        AssertViewerPreferences(PdfMerger.Merge(viewerPreferencePdf));
        AssertViewerPreferences(PdfStamper.StampText(viewerPreferencePdf, "STAMP"));

        static void AssertViewerPreferences(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/ViewerPreferences ", text, StringComparison.Ordinal);
            Assert.Contains("/HideToolbar true", text, StringComparison.Ordinal);
            Assert.Contains("/DisplayDocTitle true", text, StringComparison.Ordinal);

            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.True(info.HasViewerPreferences);
            Assert.True(info.HasReadableViewerPreferences);
            Assert.NotNull(info.ViewerPreferences);
            Assert.Equal(2, info.ViewerPreferences!.Count);
            Assert.True(info.ViewerPreferences.GetBoolean("HideToolbar"));
            Assert.True(info.ViewerPreferences.GetBoolean("DisplayDocTitle"));
        }
    }

    [Fact]
    public void RewriteApis_RejectComplexViewerPreferencesWithClearUnsupportedDiagnostic() {
        byte[] viewerPreferencePdf = BuildComplexViewerPreferencePdf();

        static void AssertViewerPreferences(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertViewerPreferences(() => PdfPageExtractor.ExtractPages(viewerPreferencePdf, 1));
        AssertViewerPreferences(() => PdfPageExtractor.SplitPages(viewerPreferencePdf));
        AssertViewerPreferences(() => PdfPageEditor.DeletePages(viewerPreferencePdf, 1));
        AssertViewerPreferences(() => PdfMetadataEditor.UpdateMetadata(viewerPreferencePdf, title: "Updated"));
        AssertViewerPreferences(() => PdfMerger.Merge(viewerPreferencePdf));
        AssertViewerPreferences(() => PdfStamper.StampText(viewerPreferencePdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_RejectTaggedPdfsWithClearUnsupportedDiagnostic() {
        byte[] taggedPdf = BuildTaggedPdf();

        AssertTagged(() => PdfPageExtractor.ExtractPages(taggedPdf, 1));
        AssertTagged(() => PdfPageExtractor.SplitPages(taggedPdf));
        AssertTagged(() => PdfPageEditor.DeletePages(taggedPdf, 1));
        AssertTagged(() => PdfMetadataEditor.UpdateMetadata(taggedPdf, title: "Updated"));
        AssertTagged(() => PdfMerger.Merge(taggedPdf));
        AssertTagged(() => PdfStamper.StampText(taggedPdf, "STAMP"));

        static void AssertTagged(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_PreserveSimpleXmpMetadata() {
        byte[] xmpMetadataPdf = BuildXmpMetadataPdf();

        AssertXmpMetadata(PdfPageExtractor.ExtractPages(xmpMetadataPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(xmpMetadataPdf);
        Assert.Single(splitPages);
        AssertXmpMetadata(splitPages[0]);
        AssertXmpMetadata(PdfPageEditor.DeletePages(BuildTwoPageXmpMetadataPdf(), 2));
        AssertXmpMetadata(PdfPageEditor.ReorderPages(BuildTwoPageXmpMetadataPdf(), 2, 1));
        AssertXmpMetadata(PdfPageEditor.RotatePages(xmpMetadataPdf, 90));
        AssertXmpMetadata(PdfMetadataEditor.UpdateMetadata(xmpMetadataPdf, title: "Updated"));
        AssertXmpMetadata(PdfMerger.Merge(xmpMetadataPdf));
        AssertXmpMetadata(PdfStamper.StampText(xmpMetadataPdf, "STAMP"));

        static void AssertXmpMetadata(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Metadata ", text, StringComparison.Ordinal);
            Assert.Contains("/Type /Metadata", text, StringComparison.Ordinal);
            Assert.Contains("/Subtype /XML", text, StringComparison.Ordinal);
            Assert.Contains("<x:xmpmeta/>", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_RejectComplexXmpMetadataWithClearUnsupportedDiagnostic() {
        byte[] xmpMetadataPdf = BuildComplexXmpMetadataPdf();

        static void AssertXmpMetadata(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertXmpMetadata(() => PdfPageExtractor.ExtractPages(xmpMetadataPdf, 1));
        AssertXmpMetadata(() => PdfPageExtractor.SplitPages(xmpMetadataPdf));
        AssertXmpMetadata(() => PdfPageEditor.DeletePages(xmpMetadataPdf, 1));
        AssertXmpMetadata(() => PdfMetadataEditor.UpdateMetadata(xmpMetadataPdf, title: "Updated"));
        AssertXmpMetadata(() => PdfMerger.Merge(xmpMetadataPdf));
        AssertXmpMetadata(() => PdfStamper.StampText(xmpMetadataPdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_PreserveSimpleOutputIntents() {
        byte[] outputIntentPdf = BuildOutputIntentPdf();

        AssertOutputIntents(PdfPageExtractor.ExtractPages(outputIntentPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(outputIntentPdf);
        Assert.Single(splitPages);
        AssertOutputIntents(splitPages[0]);
        AssertOutputIntents(PdfPageEditor.DeletePages(BuildTwoPageOutputIntentPdf(), 2));
        AssertOutputIntents(PdfPageEditor.ReorderPages(BuildTwoPageOutputIntentPdf(), 2, 1));
        AssertOutputIntents(PdfPageEditor.RotatePages(outputIntentPdf, 90));
        AssertOutputIntents(PdfMetadataEditor.UpdateMetadata(outputIntentPdf, title: "Updated"));
        AssertOutputIntents(PdfMerger.Merge(outputIntentPdf));
        AssertOutputIntents(PdfStamper.StampText(outputIntentPdf, "STAMP"));

        static void AssertOutputIntents(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/OutputIntents [", text, StringComparison.Ordinal);
            Assert.Contains("/Type /OutputIntent", text, StringComparison.Ordinal);
            Assert.Contains("/OutputConditionIdentifier (sRGB)", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_RejectComplexOutputIntentPdfsWithClearUnsupportedDiagnostic() {
        byte[] outputIntentPdf = BuildComplexOutputIntentPdf();

        static void AssertOutputIntents(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertOutputIntents(() => PdfPageExtractor.ExtractPages(outputIntentPdf, 1));
        AssertOutputIntents(() => PdfPageExtractor.SplitPages(outputIntentPdf));
        AssertOutputIntents(() => PdfPageEditor.DeletePages(outputIntentPdf, 1));
        AssertOutputIntents(() => PdfMetadataEditor.UpdateMetadata(outputIntentPdf, title: "Updated"));
        AssertOutputIntents(() => PdfMerger.Merge(outputIntentPdf));
        AssertOutputIntents(() => PdfStamper.StampText(outputIntentPdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_PreserveSimpleEmbeddedFiles() {
        byte[] embeddedFilePdf = BuildEmbeddedFilePdf();

        AssertEmbeddedFiles(PdfPageExtractor.ExtractPages(embeddedFilePdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(embeddedFilePdf);
        Assert.Single(splitPages);
        AssertEmbeddedFiles(splitPages[0]);
        AssertEmbeddedFiles(PdfPageEditor.DeletePages(BuildTwoPageEmbeddedFilePdf(), 2));
        AssertEmbeddedFiles(PdfPageEditor.ReorderPages(BuildTwoPageEmbeddedFilePdf(), 2, 1));
        AssertEmbeddedFiles(PdfPageEditor.RotatePages(embeddedFilePdf, 90));
        AssertEmbeddedFiles(PdfMetadataEditor.UpdateMetadata(embeddedFilePdf, title: "Updated"));
        AssertEmbeddedFiles(PdfMerger.Merge(embeddedFilePdf));
        AssertEmbeddedFiles(PdfStamper.StampText(embeddedFilePdf, "STAMP"));

        static void AssertEmbeddedFiles(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Names << /EmbeddedFiles", text, StringComparison.Ordinal);
            Assert.Contains("/Type /Filespec", text, StringComparison.Ordinal);
            Assert.Contains("/F (note.txt)", text, StringComparison.Ordinal);
            Assert.Contains("/Type /EmbeddedFile", text, StringComparison.Ordinal);
            Assert.Contains("note", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_PreserveSimpleAssociatedFiles() {
        byte[] associatedFilePdf = BuildAssociatedFilePdf();

        AssertAssociatedFiles(PdfPageExtractor.ExtractPages(associatedFilePdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(associatedFilePdf);
        Assert.Single(splitPages);
        AssertAssociatedFiles(splitPages[0]);
        AssertAssociatedFiles(PdfPageEditor.DeletePages(BuildTwoPageAssociatedFilePdf(), 2));
        AssertAssociatedFiles(PdfPageEditor.ReorderPages(BuildTwoPageAssociatedFilePdf(), 2, 1));
        AssertAssociatedFiles(PdfPageEditor.RotatePages(associatedFilePdf, 90));
        AssertAssociatedFiles(PdfMetadataEditor.UpdateMetadata(associatedFilePdf, title: "Updated"));
        AssertAssociatedFiles(PdfMerger.Merge(associatedFilePdf));
        AssertAssociatedFiles(PdfStamper.StampText(associatedFilePdf, "STAMP"));

        static void AssertAssociatedFiles(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/AF [", text, StringComparison.Ordinal);
            Assert.Contains("/Type /Filespec", text, StringComparison.Ordinal);
            Assert.Contains("/AFRelationship /Data", text, StringComparison.Ordinal);
            Assert.Contains("/F (data.xml)", text, StringComparison.Ordinal);
            Assert.Contains("/Type /EmbeddedFile", text, StringComparison.Ordinal);
            Assert.Contains("data", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_PreserveCombinedDestinationAndEmbeddedFileNameTrees() {
        byte[] combinedNameTreePdf = BuildCombinedDestinationAndEmbeddedFileNameTreePdf();

        AssertCombinedNameTrees(PdfPageExtractor.ExtractPages(combinedNameTreePdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(combinedNameTreePdf);
        Assert.Single(splitPages);
        AssertCombinedNameTrees(splitPages[0]);
        AssertCombinedNameTrees(PdfPageEditor.RotatePages(combinedNameTreePdf, 90));
        AssertCombinedNameTrees(PdfMetadataEditor.UpdateMetadata(combinedNameTreePdf, title: "Updated"));
        AssertCombinedNameTrees(PdfMerger.Merge(combinedNameTreePdf));
        AssertCombinedNameTrees(PdfStamper.StampText(combinedNameTreePdf, "STAMP"));

        static void AssertCombinedNameTrees(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Names << /Dests << /Names [", text, StringComparison.Ordinal);
            Assert.Contains("(Chapter1)", text, StringComparison.Ordinal);
            Assert.Contains("/XYZ 0 200 0", text, StringComparison.Ordinal);
            Assert.Contains("/EmbeddedFiles << /Names [", text, StringComparison.Ordinal);
            Assert.Contains("(note.txt)", text, StringComparison.Ordinal);
            Assert.Contains("/Type /Filespec", text, StringComparison.Ordinal);
            Assert.Contains("/Type /EmbeddedFile", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_RejectUnsupportedCatalogNameTreesWithClearUnsupportedDiagnostic() {
        byte[] nameTreePdf = BuildUnsupportedCatalogNameTreePdf();

        static void AssertCatalogNameTrees(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("catalog name trees", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        AssertCatalogNameTrees(() => PdfPageExtractor.ExtractPages(nameTreePdf, 1));
        AssertCatalogNameTrees(() => PdfPageExtractor.SplitPages(nameTreePdf));
        AssertCatalogNameTrees(() => PdfPageEditor.DeletePages(nameTreePdf, 1));
        AssertCatalogNameTrees(() => PdfMetadataEditor.UpdateMetadata(nameTreePdf, title: "Updated"));
        AssertCatalogNameTrees(() => PdfMerger.Merge(nameTreePdf));
        AssertCatalogNameTrees(() => PdfStamper.StampText(nameTreePdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_RejectComplexEmbeddedFilePdfsWithClearUnsupportedDiagnostic() {
        byte[] embeddedFilePdf = BuildComplexEmbeddedFilePdf();

        static void AssertEmbeddedFiles(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertEmbeddedFiles(() => PdfPageExtractor.ExtractPages(embeddedFilePdf, 1));
        AssertEmbeddedFiles(() => PdfPageExtractor.SplitPages(embeddedFilePdf));
        AssertEmbeddedFiles(() => PdfPageEditor.DeletePages(embeddedFilePdf, 1));
        AssertEmbeddedFiles(() => PdfMetadataEditor.UpdateMetadata(embeddedFilePdf, title: "Updated"));
        AssertEmbeddedFiles(() => PdfMerger.Merge(embeddedFilePdf));
        AssertEmbeddedFiles(() => PdfStamper.StampText(embeddedFilePdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_RejectComplexAssociatedFilePdfsWithClearUnsupportedDiagnostic() {
        byte[] associatedFilePdf = BuildComplexAssociatedFilePdf();

        static void AssertEmbeddedFiles(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertEmbeddedFiles(() => PdfPageExtractor.ExtractPages(associatedFilePdf, 1));
        AssertEmbeddedFiles(() => PdfPageExtractor.SplitPages(associatedFilePdf));
        AssertEmbeddedFiles(() => PdfPageEditor.DeletePages(associatedFilePdf, 1));
        AssertEmbeddedFiles(() => PdfMetadataEditor.UpdateMetadata(associatedFilePdf, title: "Updated"));
        AssertEmbeddedFiles(() => PdfMerger.Merge(associatedFilePdf));
        AssertEmbeddedFiles(() => PdfStamper.StampText(associatedFilePdf, "STAMP"));
    }


    [Fact]
    public void RewriteApis_PreserveSimpleOptionalContent() {
        byte[] optionalContentPdf = BuildOptionalContentPdf();

        AssertOptionalContent(PdfPageExtractor.ExtractPages(optionalContentPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(optionalContentPdf);
        Assert.Single(splitPages);
        AssertOptionalContent(splitPages[0]);
        AssertOptionalContent(PdfPageEditor.DeletePages(BuildTwoPageOptionalContentPdf(), 2));
        AssertOptionalContent(PdfPageEditor.ReorderPages(BuildTwoPageOptionalContentPdf(), 2, 1));
        AssertOptionalContent(PdfPageEditor.RotatePages(optionalContentPdf, 90));
        AssertOptionalContent(PdfMetadataEditor.UpdateMetadata(optionalContentPdf, title: "Updated"));
        AssertOptionalContent(PdfMerger.Merge(optionalContentPdf));
        AssertOptionalContent(PdfStamper.StampText(optionalContentPdf, "STAMP"));

        static void AssertOptionalContent(byte[] output) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/OCProperties ", text, StringComparison.Ordinal);
            Assert.Contains("/OCGs [", text, StringComparison.Ordinal);
            Assert.Contains("/Type /OCG", text, StringComparison.Ordinal);
            Assert.Contains("/Name (Layer 1)", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_RejectComplexOptionalContentPdfsWithClearUnsupportedDiagnostic() {
        byte[] optionalContentPdf = BuildComplexOptionalContentPdf();

        static void AssertOptionalContent(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertOptionalContent(() => PdfPageExtractor.ExtractPages(optionalContentPdf, 1));
        AssertOptionalContent(() => PdfPageExtractor.SplitPages(optionalContentPdf));
        AssertOptionalContent(() => PdfPageEditor.DeletePages(optionalContentPdf, 1));
        AssertOptionalContent(() => PdfMetadataEditor.UpdateMetadata(optionalContentPdf, title: "Updated"));
        AssertOptionalContent(() => PdfMerger.Merge(optionalContentPdf));
        AssertOptionalContent(() => PdfStamper.StampText(optionalContentPdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_RejectActiveContentPdfsWithClearUnsupportedDiagnostic() {
        byte[] activeContentPdf = BuildActiveContentPdf();

        AssertActiveContent(() => PdfPageExtractor.ExtractPages(activeContentPdf, 1));
        AssertActiveContent(() => PdfPageExtractor.SplitPages(activeContentPdf));
        AssertActiveContent(() => PdfPageEditor.DeletePages(activeContentPdf, 1));
        AssertActiveContent(() => PdfMetadataEditor.UpdateMetadata(activeContentPdf, title: "Updated"));
        AssertActiveContent(() => PdfMerger.Merge(activeContentPdf));
        AssertActiveContent(() => PdfStamper.StampText(activeContentPdf, "STAMP"));

        static void AssertActiveContent(Action action) {
            var exception = Assert.Throws<NotSupportedException>(action);
            Assert.Contains("PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }
    }

    private static byte[] BuildPdf() {
        return PdfDoc.Create()
            .Meta(title: "Stream read", author: "OfficeIMO", subject: "Read", keywords: "stream,pdf")
            .Paragraph(p => p.Text("Stream readable text"))
            .ToBytes();
    }

    private static byte[] BuildTwoPageLinkAnnotationPdf() {
        return PdfDoc.Create()
            .Paragraph(p => p.Link("First", "https://evotec.xyz/first", contents: "First link metadata"))
            .PageBreak()
            .Paragraph(p => p.Link("Second", "https://evotec.xyz/second", contents: "Second link metadata"))
            .ToBytes();
    }

    private static byte[] BuildOutlinePdf() {
        return PdfDoc.Create(new PdfOptions { CreateOutlineFromHeadings = true })
            .H1("Executive summary")
            .Paragraph(p => p.Text("Outline sample"))
            .ToBytes();
    }

    private static byte[] BuildTwoPageOutlinePdf() {
        return PdfDoc.Create(new PdfOptions { CreateOutlineFromHeadings = true })
            .H1("Executive summary")
            .Paragraph(p => p.Text("Outline sample"))
            .PageBreak()
            .H1("Appendix")
            .Paragraph(p => p.Text("Appendix sample"))
            .ToBytes();
    }

    private static byte[] BuildGoToActionOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Chapter 1) /Parent 5 0 R /A << /S /GoTo /D [3 0 R /XYZ 0 200 0] >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildGoToActionIndirectDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Indirect GoTo action) /Parent 5 0 R /A << /S /GoTo /D 7 0 R >> >>",
            "endobj",
            "7 0 obj",
            "[3 0 R /XYZ 0 188 0]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildGoToActionDictionaryDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Dictionary GoTo action) /Parent 5 0 R /A << /S /GoTo /D << /D [3 0 R /XYZ 0 188 0] >> >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCyclicGoToActionDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Cyclic GoTo action) /Parent 5 0 R /A << /S /GoTo /D 7 0 R >> >>",
            "endobj",
            "7 0 obj",
            "8 0 R",
            "endobj",
            "8 0 obj",
            "7 0 R",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildUriActionOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (External) /Parent 5 0 R /A << /S /URI /URI (https://evotec.xyz) >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildIndirectOutlineDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Indirect destination) /Parent 5 0 R /Dest 7 0 R >>",
            "endobj",
            "7 0 obj",
            "[3 0 R /XYZ 0 144 0]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildStaleCatalogRevisionPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLayout /TwoColumnLeft /PageLabels 6 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLayout /SinglePage >>",
            "endobj",
            "6 0 obj",
            "<< /Kids [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /Nums [0 << /S /D /St 10 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 5 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildStaleCatalogWithDifferentPagesAndOutlinesPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 4 0 R /PageLayout /TwoColumnLeft >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 11 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Outlines /First 12 0 R /Last 12 0 R /Count 1 >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Catalog /Pages 6 0 R /Outlines 8 0 R /PageLayout /SinglePage >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Pages /Count 1 /Kids [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Page /Parent 6 0 R /MediaBox [0 0 200 200] /Contents 11 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Outlines /First 9 0 R /Last 9 0 R /Count 1 >>",
            "endobj",
            "9 0 obj",
            "<< /Title (Current) /Parent 8 0 R /Dest [7 0 R /XYZ 0 144 0] >>",
            "endobj",
            "11 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "12 0 obj",
            "<< /Title (Old) /Parent 4 0 R /Dest [3 0 R /XYZ 0 72 0] >>",
            "endobj",
            "trailer",
            "<< /Root 5 0 R /Size 13 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildStaleCatalogPageLabelsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLabels 13 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 11 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 11 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Catalog /Pages 6 0 R /PageLabels 10 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Pages /Count 2 /Kids [7 0 R 8 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Page /Parent 6 0 R /MediaBox [0 0 200 200] /Contents 11 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Page /Parent 6 0 R /MediaBox [0 0 200 200] /Contents 11 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /Nums [0 << /S /r /St 1 >> 1 << /S /D /St 10 >>] >>",
            "endobj",
            "11 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "13 0 obj",
            "<< /Nums [0 << /S /A /St 1 >> 1 << /S /A /St 2 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 5 0 R /Size 14 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCatalogViewSettingPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageMode /FullScreen /PageLayout /TwoColumnLeft >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageCatalogViewSettingPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageMode /FullScreen /PageLayout /TwoColumnLeft >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCatalogIdentityPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Version /1.7 /Lang (pl-PL) >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageCatalogIdentityPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Version /1.7 /Lang (pl-PL) >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCatalogUriPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /URI << /Base (https://evotec.xyz/docs/) >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageCatalogUriPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /URI << /Base (https://evotec.xyz/docs/) >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexCatalogUriPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /URI << /Base (https://evotec.xyz/docs/) /SourcePage 3 0 R >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildEncryptedPdfMarker() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [] >>",
            "endobj",
            "5 0 obj",
            "<< /Filter /Standard /V 1 /R 2 /O () /U () /P -4 >>",
            "endobj",
            "xref",
            "0 6",
            "0000000000 65535 f ",
            "trailer",
            "<< /Root 1 0 R /Size 6 /Encrypt 5 0 R >>",
            "startxref",
            "0",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSignedPdfMarker() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Sig /ByteRange [0 0 0 0] /Contents <> >>",
            "endobj",
            "6 0 obj",
            "<< /SigFlags 3 /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Sig /V 5 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFormPdfMarker() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [6 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /FT /Tx /T (Name) /V (OfficeIMO) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPageLabelPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLabels 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Nums [0 << /S /D /St 1 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageLabelPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLabels 7 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Nums [0 << /S /D /St 1 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexPageLabelPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLabels 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Kids [6 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /Nums [0 << /S /D /St 1 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNamedDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Dests 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Chapter1 [3 0 R /XYZ 0 200 0] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageNamedDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Dests 7 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Chapter1 [3 0 R /XYZ 0 200 0] /Chapter2 [5 0 R /Fit] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNamedDestinationNameTreePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Names [(Chapter1) [3 0 R /XYZ 0 200 0]] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageNamedDestinationNameTreePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Names [(Chapter1) [3 0 R /XYZ 0 200 0] (Chapter2) [5 0 R /Fit]] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNamedDestinationNameTreeWithKidsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Kids [5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Names [(Chapter1) [3 0 R /XYZ 0 200 0]] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageNamedDestinationNameTreeWithKidsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Kids [7 0 R 8 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Names [(Chapter1) [3 0 R /XYZ 0 200 0]] >>",
            "endobj",
            "8 0 obj",
            "<< /Names [(Chapter2) [5 0 R /Fit]] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexNamedDestinationNameTreePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Kids [5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Names [(Chapter1) [6 0 R /XYZ 0 200 0]] >>",
            "endobj",
            "6 0 obj",
            "<< /NotAPage true >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOpenActionPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "[3 0 R /Fit]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageOpenActionPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction 7 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "[5 0 R /Fit]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOpenActionDictionaryPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /GoTo /D [3 0 R /Fit] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageOpenActionDictionaryPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction 7 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /S /GoTo /D [5 0 R /Fit] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexOpenActionDictionaryPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /URI /URI (https://example.com) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildViewerPreferencePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /ViewerPreferences 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /HideToolbar true /DisplayDocTitle true >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageViewerPreferencePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /ViewerPreferences 7 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /HideToolbar true /DisplayDocTitle true >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexViewerPreferencePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /ViewerPreferences 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /HideToolbar true /ViewArea 3 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTaggedPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /StructTreeRoot 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /StructParents 0 >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /StructTreeRoot /K [6 0 R] /ParentTree 7 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Type /StructElem /S /Document /P 5 0 R >>",
            "endobj",
            "7 0 obj",
            "<< /Nums [0 [6 0 R]] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildXmpMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Metadata /Subtype /XML /Length 12 >>",
            "stream",
            "<x:xmpmeta/>",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageXmpMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Metadata 7 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Metadata /Subtype /XML /Length 12 >>",
            "stream",
            "<x:xmpmeta/>",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexXmpMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Metadata /Subtype /XML /Source 3 0 R /Length 12 >>",
            "stream",
            "<x:xmpmeta/>",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOutputIntentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OutputIntents [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageOutputIntentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OutputIntents [7 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexOutputIntentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OutputIntents [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB) /DestOutputProfile 3 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildEmbeddedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(note.txt) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (note.txt) /EF << /F 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /EmbeddedFile /Length 4 >>",
            "stream",
            "note",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageEmbeddedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(note.txt) 7 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Filespec /F (note.txt) /EF << /F 8 0 R >> >>",
            "endobj",
            "8 0 obj",
            "<< /Type /EmbeddedFile /Length 4 >>",
            "stream",
            "note",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildAssociatedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AF [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (data.xml) /AFRelationship /Data /EF << /F 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /EmbeddedFile /Subtype /text#2Fxml /Length 4 >>",
            "stream",
            "data",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageAssociatedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AF [7 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Filespec /F (data.xml) /AFRelationship /Data /EF << /F 8 0 R >> >>",
            "endobj",
            "8 0 obj",
            "<< /Type /EmbeddedFile /Subtype /text#2Fxml /Length 4 >>",
            "stream",
            "data",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCombinedDestinationAndEmbeddedFileNameTreePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Names [(Chapter1) [3 0 R /XYZ 0 200 0]] >> /EmbeddedFiles << /Names [(note.txt) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (note.txt) /EF << /F 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /EmbeddedFile /Length 4 >>",
            "stream",
            "note",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildUnsupportedCatalogNameTreePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Templates << /Names [(Layout) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Template /Name (Layout) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexEmbeddedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(note.txt) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (note.txt) /EF << /F 3 0 R >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexAssociatedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AF [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (data.xml) /AFRelationship /Data /EF << /F 3 0 R >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOptionalContentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.5",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [5 0 R] /D << /ON [5 0 R] /Order [5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /OCG /Name (Layer 1) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageOptionalContentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.5",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [7 0 R] /D << /ON [7 0 R] /Order [7 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /OCG /Name (Layer 1) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexOptionalContentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.5",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [3 0 R] /D << /ON [3 0 R] /Order [3 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildActiveContentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Open) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static MemoryStream BuildPrefixedStream(byte[] pdf) {
        var data = new byte[pdf.Length + 5];
        data[0] = 1;
        data[1] = 2;
        data[2] = 3;
        data[3] = 4;
        data[4] = 5;
        Array.Copy(pdf, 0, data, 5, pdf.Length);
        return new MemoryStream(data);
    }

    private sealed class WriteOnlyStream : Stream {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => 0;

        public override long Position {
            get => 0;
            set => throw new NotSupportedException();
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin) {
            throw new NotSupportedException();
        }

        public override void SetLength(long value) {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count) {
        }
    }
}
