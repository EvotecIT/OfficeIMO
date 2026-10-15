using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
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
            var attachment = Assert.Single(PdfReadDocument.Open(output).ExtractAttachments());
            Assert.Equal("note.txt", attachment.FileName);
            Assert.Equal("note", System.Text.Encoding.ASCII.GetString(attachment.Bytes));
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
            var attachment = Assert.Single(PdfReadDocument.Open(output).ExtractAttachments());
            Assert.Equal("data.xml", attachment.FileName);
            Assert.Equal(PdfAssociatedFileRelationship.Data, attachment.Relationship);
            Assert.Equal("AF", attachment.Source);
            Assert.Equal("data", System.Text.Encoding.ASCII.GetString(attachment.Bytes));
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
            PdfReadDocument document = PdfReadDocument.Open(output);
            PdfNamedDestination destination = Assert.Single(document.NamedDestinations);
            Assert.Equal("Chapter1", destination.Name);
            Assert.Equal(1, destination.PageNumber);
            Assert.Equal(PdfOpenActionDestinationMode.Xyz, destination.DestinationMode);
            Assert.Equal(0, destination.DestinationLeft);
            Assert.Equal(200, destination.DestinationTop);
            Assert.Equal(0, destination.DestinationZoom);
            var attachment = Assert.Single(document.ExtractAttachments());
            Assert.Equal("note.txt", attachment.FileName);
            Assert.Equal("note", System.Text.Encoding.ASCII.GetString(attachment.Bytes));
        }
    }

    [Fact]
    public void RewriteApis_RejectUnsupportedCatalogNameTreesWithClearUnsupportedDiagnostic() {
        byte[] nameTreePdf = BuildUnsupportedCatalogNameTreePdf();

        static void AssertCatalogNameTrees(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("catalog name trees", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        AssertCatalogNameTrees(() => PdfPageExtractor.ExtractPages(nameTreePdf, 1));
        AssertCatalogNameTrees(() => PdfPageExtractor.SplitPages(nameTreePdf));
        AssertCatalogNameTrees(() => PdfMetadataEditor.UpdateMetadata(nameTreePdf, title: "Updated"));
        AssertCatalogNameTrees(() => PdfMerger.Merge(nameTreePdf));
        AssertCatalogNameTrees(() => PdfStamper.StampText(nameTreePdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_RejectComplexEmbeddedFilePdfsWithClearUnsupportedDiagnostic() {
        byte[] embeddedFilePdf = BuildComplexEmbeddedFilePdf();

        static void AssertEmbeddedFiles(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertEmbeddedFiles(() => PdfPageExtractor.ExtractPages(embeddedFilePdf, 1));
        AssertEmbeddedFiles(() => PdfPageExtractor.SplitPages(embeddedFilePdf));
        AssertEmbeddedFiles(() => PdfMetadataEditor.UpdateMetadata(embeddedFilePdf, title: "Updated"));
        AssertEmbeddedFiles(() => PdfMerger.Merge(embeddedFilePdf));
        AssertEmbeddedFiles(() => PdfStamper.StampText(embeddedFilePdf, "STAMP"));
    }

    [Fact]
    public void RewriteApis_RejectComplexAssociatedFilePdfsWithClearUnsupportedDiagnostic() {
        byte[] associatedFilePdf = BuildComplexAssociatedFilePdf();

        static void AssertEmbeddedFiles(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertEmbeddedFiles(() => PdfPageExtractor.ExtractPages(associatedFilePdf, 1));
        AssertEmbeddedFiles(() => PdfPageExtractor.SplitPages(associatedFilePdf));
        AssertEmbeddedFiles(() => PdfMetadataEditor.UpdateMetadata(associatedFilePdf, title: "Updated"));
        AssertEmbeddedFiles(() => PdfMerger.Merge(associatedFilePdf));
        AssertEmbeddedFiles(() => PdfStamper.StampText(associatedFilePdf, "STAMP"));
    }



}
