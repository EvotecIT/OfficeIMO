using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
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
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
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
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }
    }


}
