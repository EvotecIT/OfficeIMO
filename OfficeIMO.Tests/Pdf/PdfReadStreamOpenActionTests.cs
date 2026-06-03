using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
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


}
