using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
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
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertViewerPreferences(() => PdfPageExtractor.ExtractPages(viewerPreferencePdf, 1));
        AssertViewerPreferences(() => PdfPageExtractor.SplitPages(viewerPreferencePdf));
        AssertViewerPreferences(() => PdfMetadataEditor.UpdateMetadata(viewerPreferencePdf, title: "Updated"));
        AssertViewerPreferences(() => PdfMerger.Merge(viewerPreferencePdf));
        AssertViewerPreferences(() => PdfStamper.StampText(viewerPreferencePdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_ApplyOperationSpecificTaggedStructureGates() {
        byte[] taggedPdf = BuildTaggedPdf();

        Assert.True(PdfInspector.Inspect(PdfPageExtractor.ExtractPages(taggedPdf, 1)).HasTaggedContent);
        Assert.True(PdfInspector.Inspect(Assert.Single(PdfPageExtractor.SplitPages(taggedPdf))).HasTaggedContent);
        AssertTagged(() => PdfMetadataEditor.UpdateMetadata(taggedPdf, title: "Updated"));
        Assert.True(PdfInspector.Inspect(PdfMerger.Merge(taggedPdf)).HasTaggedContent);
        AssertTagged(() => PdfStamper.StampText(taggedPdf, "STAMP"));

        static void AssertTagged(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
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
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertXmpMetadata(() => PdfPageExtractor.ExtractPages(xmpMetadataPdf, 1));
        AssertXmpMetadata(() => PdfPageExtractor.SplitPages(xmpMetadataPdf));
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
            var intent = Assert.Single(PdfInspector.Inspect(output).OutputIntents);
            Assert.Equal("sRGB", intent.OutputConditionIdentifier);
        }
    }

    [Fact]
    public void RewriteApis_RejectComplexOutputIntentPdfsWithClearUnsupportedDiagnostic() {
        byte[] outputIntentPdf = BuildComplexOutputIntentPdf();

        static void AssertOutputIntents(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertOutputIntents(() => PdfPageExtractor.ExtractPages(outputIntentPdf, 1));
        AssertOutputIntents(() => PdfPageExtractor.SplitPages(outputIntentPdf));
        AssertOutputIntents(() => PdfMetadataEditor.UpdateMetadata(outputIntentPdf, title: "Updated"));
        AssertOutputIntents(() => PdfMerger.Merge(outputIntentPdf));
        AssertOutputIntents(() => PdfStamper.StampText(outputIntentPdf, "STAMP"));
    }


}
