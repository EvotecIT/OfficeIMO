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
            var group = Assert.Single(PdfInspector.Inspect(output).OptionalContentGroups);
            Assert.Equal("Layer 1", group.Name);
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
        AssertOptionalContent(() => PdfMetadataEditor.UpdateMetadata(optionalContentPdf, title: "Updated"));
        AssertOptionalContent(() => PdfMerger.Merge(optionalContentPdf));
        AssertOptionalContent(() => PdfStamper.StampText(optionalContentPdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_RejectActiveContentPdfsWithClearUnsupportedDiagnostic() {
        byte[] activeContentPdf = BuildActiveContentPdf();

        AssertActiveContent(() => PdfPageExtractor.ExtractPages(activeContentPdf, 1));
        AssertActiveContent(() => PdfPageExtractor.SplitPages(activeContentPdf));
        AssertActiveContent(() => PdfMetadataEditor.UpdateMetadata(activeContentPdf, title: "Updated"));
        AssertActiveContent(() => PdfMerger.Merge(activeContentPdf));
        AssertActiveContent(() => PdfStamper.StampText(activeContentPdf, "STAMP"));

        static void AssertActiveContent(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }
    }


}
