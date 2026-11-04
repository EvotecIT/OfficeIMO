using System;
using System.IO;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;
using static OfficeIMO.Xps.Tests.XpsLogicalStructureTests;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsPdfStreamTests {
    [Fact]
    public void RepeatedPageOccurrencesShareProjectionCommandBudget() {
        var document = XpsDocument.Create();
        var page = document.AddPage(100, 100).AddPath("M10,10H20V20H10Z");
        var options = new XpsPdfExportOptions { MaximumPathCommands = 8 };
        using (var single = new MemoryStream())
            Assert.Equal(1, document.SaveAsPdf(single, options).ItemsWritten);

        document.Documents[0].InsertPage(1, page);
        using var output = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        output.Position = 2;

        InvalidDataException error = Assert.Throws<InvalidDataException>(() => document.SaveAsPdf(output, options));
        Assert.Contains("MaximumPathCommands", error.Message, StringComparison.Ordinal);
        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray());
        Assert.Equal(2, output.Position);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void StreamExportRetainsNativeTextAndLeavesTheCallerStreamOpen(bool semantics) {
        var document = Create(XpsFormat.OpenXps, new[] { "Alpha", "Beta" });
        document.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(document, null, null, Paragraph(document.StructureNamespace, "Beta", "Alpha")));
        using var output = new MemoryStream();
        OfficeConversionReport saved = document.SaveAsPdf(output, new XpsPdfExportOptions { PreserveLogicalStructure = semantics });
        Assert.Equal(document.Pages.Count, saved.ItemsWritten);
        Assert.True(output.CanWrite); Assert.True(output.Length > 0);
        var read = PdfDocument.Load(new System.IO.MemoryStream(output.ToArray()));
        Assert.Equal(semantics, read.TaggedContent != null);
        Assert.Contains("Alpha", PdfDocument.Load(new System.IO.MemoryStream(output.ToArray())).Read().Text);
        Assert.Contains("Beta", PdfDocument.Load(new System.IO.MemoryStream(output.ToArray())).Read().Text);
    }

    [Fact]
    public void StreamExportAppliesTheClonedNativePdfSecuritySettings() {
        var options = new XpsPdfExportOptions { PdfOptions = new PdfCreateOptions().SetEncryption("original") };
        var snapshot = options.Clone(); options.PdfOptions.SetEncryption("changed");
        using var output = new MemoryStream();
        Create(XpsFormat.Xps, new[] { "Protected" }).SaveAsPdf(output, snapshot);
        var loaded = PdfDocument.Load(new System.IO.MemoryStream(output.ToArray()), new PdfLoadOptions { Password = "original" });
        Assert.Contains("Protected", loaded.Read().Text);
    }

    [Fact]
    public void PreCanceledExportDoesNotTouchCallerOutput() {
        using var output = new MemoryStream(new byte[] { 1, 2, 3 }, true); output.Position = 2;
        Assert.Throws<OperationCanceledException>(() => Create(XpsFormat.Xps, new[] { "Text" })
            .SaveAsPdf(output, cancellationToken: new CancellationToken(true)));
        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray()); Assert.Equal(2, output.Position); Assert.True(output.CanWrite);
    }
}
