using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {

    [Fact]
    public void ExtractText_UsesXrefStreamOffsetsInsteadOfTrailingStaleDuplicateObjects() {
        byte[] pdf = BuildXrefStreamPdfWithTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Active xref stream page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_UsesXrefStreamCompressedObjectEntriesInsteadOfTrailingStaleDuplicates() {
        byte[] pdf = BuildXrefStreamCompressedObjectPdfWithTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Active compressed xref page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale compressed trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_IgnoresTrailingStaleObjectStreamsOutsideActiveXrefChain() {
        byte[] pdf = BuildXrefStreamPdfWithTrailingStaleObjectStreamPage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Active xref stream page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale trailing object stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_IgnoresTrailingStaleObjectStreamsOutsideActiveClassicXref() {
        byte[] pdf = BuildClassicXrefPdfWithTrailingStaleObjectStreamPage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Active classic xref page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale classic object stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_IgnoresClassicXrefEntryWhenObjectGenerationDoesNotMatch() {
        byte[] pdf = BuildIncrementalClassicXrefPdfWithWrongGenerationReplacementPage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Previous classic generation page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Wrong classic generation page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_IgnoresXrefStreamEntryWhenObjectGenerationDoesNotMatch() {
        byte[] pdf = BuildIncrementalXrefStreamPdfWithWrongGenerationReplacementPage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Previous xref generation page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Wrong xref generation page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_DoesNotResolveContentReferenceWithWrongGeneration() {
        byte[] pdf = BuildClassicXrefPdfWithWrongGenerationContentReference();

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPage(pdf);

        string pageText = Normalize(Assert.Single(pages));
        Assert.DoesNotContain("Wrong generation referenced content", pageText, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_FollowsXrefStreamPrevChainForInheritedObjects() {
        byte[] pdf = BuildIncrementalXrefStreamPdfWithTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Inherited previous xref page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale incremental trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_FollowsXrefStreamPrevToClassicXrefForInheritedObjects() {
        byte[] pdf = BuildIncrementalXrefStreamPdfWithClassicPrevAndTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Inherited mixed xref page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale mixed xref trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_HonorsXrefStreamFreeEntriesOverTrailingStaleObjects() {
        byte[] pdf = BuildIncrementalXrefStreamPdfWithFreedTrailingStalePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Active replacement xref stream page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Superseded xref stream page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale freed xref stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_FollowsClassicXrefPrevChainForInheritedObjects() {
        byte[] pdf = BuildIncrementalClassicXrefPdfWithTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Inherited classic xref page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale classic trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_UsesHybridClassicXrefStmSupplementInsteadOfTrailingStaleDuplicate() {
        byte[] pdf = BuildHybridClassicXrefPdfWithXRefStmAndTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Hybrid xref stream page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale hybrid trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Inspect_UsesHybridClassicXrefStmTrailerRootBeforeStaleXrefStreamRoot() {
        byte[] pdf = BuildHybridClassicXrefPdfWithXRefStmTrailerRootAndStaleXrefStreamRoot();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal("SinglePage", info.CatalogPageLayout);
        Assert.Equal(200d, page.Width);
        Assert.Equal(200d, page.Height);
    }

    [Fact]
    public void Inspect_FollowsClassicTrailerPrevChainForInheritedRoot() {
        byte[] pdf = BuildIncrementalClassicXrefPdfWithInheritedTrailerRoot();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal("SinglePage", info.CatalogPageLayout);
        Assert.Equal(200d, page.Width);
        Assert.Equal(200d, page.Height);
    }

    [Fact]
    public void Inspect_FollowsActiveXrefStreamTrailerPrevChainForInheritedRoot() {
        byte[] pdf = BuildIncrementalXrefStreamPdfWithInheritedTrailerRootAndStaleHighObjectXref();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal("SinglePage", info.CatalogPageLayout);
        Assert.Equal(200d, page.Width);
        Assert.Equal(200d, page.Height);
    }

    [Fact]
    public void Inspect_FollowsXrefStreamPrevToClassicTrailerForInheritedRoot() {
        byte[] pdf = BuildIncrementalXrefStreamPdfWithClassicTrailerRoot();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal("SinglePage", info.CatalogPageLayout);
        Assert.Equal(200d, page.Width);
        Assert.Equal(200d, page.Height);
    }

}
