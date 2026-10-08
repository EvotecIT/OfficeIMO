using System;
using System.IO;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;
using static OfficeIMO.Xps.Tests.XpsLogicalStructureTests;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsPdfStreamTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void StreamExportRetainsNativeTextAndReplacesCallerOwnedContents(bool semantics) {
        var document = Create(XpsFormat.OpenXps, new[] { "Alpha", "Beta" });
        document.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(document, null, null, Paragraph(document.StructureNamespace, "Beta", "Alpha")));
        using var output = new MemoryStream(); output.SetLength(1024 * 1024); output.Position = 123;
        var saved = document.SavePdf(output, new XpsToPdfOptions { PreserveLogicalStructure = semantics });
        Assert.True(saved.Succeeded); Assert.Equal(output.Length, saved.BytesWritten);
        Assert.Equal(0, output.Position); Assert.True(output.CanWrite); Assert.InRange(output.Length, 1, 1024 * 1024 - 1);
        var read = PdfReadDocument.Open(output.ToArray());
        Assert.Equal(semantics, read.TaggedContent != null);
        Assert.Contains("Alpha", PdfDocument.Load(output.ToArray()).Read().Text);
        Assert.Contains("Beta", PdfDocument.Load(output.ToArray()).Read().Text);
    }

    [Fact]
    public void StreamExportAppliesTheClonedNativePdfSecuritySettings() {
        var options = new XpsToPdfOptions { PdfOptions = new PdfOptions().SetEncryption("original") };
        var snapshot = options.Clone(); options.PdfOptions.SetEncryption("changed");
        using var output = new MemoryStream();
        Create(XpsFormat.Xps, new[] { "Protected" }).SavePdf(output, snapshot);
        var loaded = PdfDocument.Load(output.ToArray(), new PdfLoadOptions { Password = "original" });
        Assert.Contains("Protected", loaded.Read().Text);
    }

    [Fact]
    public void PreCanceledExportDoesNotTouchCallerOutput() {
        using var output = new MemoryStream(new byte[] { 1, 2, 3 }, true); output.Position = 2;
        Assert.Throws<OperationCanceledException>(() => Create(XpsFormat.Xps, new[] { "Text" })
            .SavePdf(output, cancellationToken: new CancellationToken(true)));
        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray()); Assert.Equal(2, output.Position); Assert.True(output.CanWrite);
    }
}
