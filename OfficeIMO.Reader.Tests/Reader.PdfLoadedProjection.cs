using OfficeIMO.Pdf;
using System.Threading;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfLoadedProjectionTests {
    [Fact]
    public void LoadedLogicalResultUsesFullProjectionByDefaultWithoutOpeningSourcePath() {
        PdfDocumentReadResult logical = PdfDocument.Load(PdfInspectionProjectionFixture.Build()).Read();
        OfficeDocumentReadResult result = logical.ReadDocument("not-an-input-file.pdf");
        Assert.Equal("not-an-input-file.pdf", result.Source.Path);
        Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("Native page text"));
        Assert.NotNull(result.Markdown);
        Assert.Single(result.Assets);
        Assert.NotEmpty(result.Forms);
        Assert.DoesNotContain("officeimo.reader.projection.text-and-metadata", result.CapabilitiesUsed);
        Assert.NotEmpty(logical.Images);
    }

    [Fact]
    public void LoadedReducedResultRetainsTextAndEvidenceWithoutReparsingOptionalImages() {
        PdfDocument pdf = PdfDocument.Load(PdfInspectionProjectionFixture.Build(32),
            new PdfLoadOptions { Limits = new PdfReadLimits { MaxContentOperands = 20 } });
        Assert.Throws<PdfReadLimitException>(() => pdf.Read());
        PdfDocumentReadResult logical = pdf.Read(new PdfLoadOptions { IncludeImages = false, IncludeFormWidgets = false });
        OfficeDocumentReadResult result = logical.ReadDocument("loaded-lean.pdf", new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, Inspection = new ReaderInspectionOptions(),
            ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 }
        });
        Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("Native page text"));
        Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("Nested form text"));
        Assert.Null(result.Markdown); Assert.Empty(result.Assets); Assert.Empty(result.Forms);
        Assert.Empty(result.Visuals); Assert.Empty(result.OcrCandidates);
        Assert.Contains("officeimo.reader.projection.text-and-metadata", result.CapabilitiesUsed);
        Assert.Contains(result.Metadata, item => item.Category == "pdf.text-decoding");
        Assert.Contains(result.Metadata, item => item.Name == "NativePageText" && item.Value == "existing-projection");
        Assert.DoesNotContain(result.Metadata, item => item.Category == "pdf.page-outcome");
    }

    [Fact]
    public void LoadedProjectionCompletesSharedBudgetsAndRestoresScopeAfterFailure() {
        PdfDocumentReadResult logical = PdfDocument.Load(PdfInspectionProjectionFixture.Build()).Read();
        Assert.Throws<ReaderResourceLimitException>(() => logical.ReadDocument(readerOptions: new ReaderOptions {
            ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 }
        }));
        OfficeDocumentReadResult lean = logical.ReadDocument(readerOptions: new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 }
        });
        Assert.Empty(lean.Assets);
        Assert.Single(logical.ReadDocument().Assets);
        Assert.Throws<ReaderResourceLimitException>(() => logical.ReadDocument(readerOptions: new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, ResourceLimits = new ReaderResourceLimits { MaxChunks = 1 }
        }, pdfOptions: new ReaderPdfOptions { ReadOptions = new PdfLoadOptions { PageSelection = PdfPageSelection.From(1, 1) } }));
        Assert.Throws<ReaderResourceLimitException>(() => logical.ReadDocument(readerOptions: new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, Inspection = new ReaderInspectionOptions { MaxRecords = 1 }
        }));
    }

    [Fact]
    public void LoadedProjectionUsesOptionNormalizationAndTerminalCancellation() {
        PdfDocumentReadResult logical = PdfDocument.Load(PdfInspectionProjectionFixture.Build()).Read();
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.ReadDocument(readerOptions: new ReaderOptions {
            Projection = (ReaderProjection)99
        }));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => logical.ReadDocument(cancellationToken: cancellation.Token));
        Assert.Single(logical.ReadDocument().Pages);
    }
}
