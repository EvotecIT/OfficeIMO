using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDebuggerTests {
    [Fact]
    public void Dump_ProjectsObjectsRevisionsResourcesOperatorsAndReachability() {
        byte[] source = PdfDocument.Create()
            .Meta(title: "Debugger original")
            .Paragraph(paragraph => paragraph.Text("Debugger body"))
            .ToBytes();
        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(source, title: "Debugger updated");

        PdfDebuggerReport report = PdfDocument.Open(updated).Debug(new PdfDebuggerOptions {
            IncludeDecodedStreamPreviews = true
        });
        PdfDebugPage page = Assert.Single(report.Pages);
        string text = report.ToText();

        Assert.True(report.Objects.Count >= 6);
        Assert.True(report.Revisions.Count >= 2);
        Assert.Contains(report.Objects, item => item.Kind == "Dictionary.Catalog" && item.Reachable);
        Assert.Contains(report.Objects, item => !item.Reachable);
        Assert.Contains("Font", page.ResourceCategories);
        Assert.NotEmpty(page.ContentObjectNumbers);
        Assert.Contains("BT", page.ContentOperators);
        Assert.Contains("Tf", page.ContentOperators);
        Assert.Contains("Tj", page.ContentOperators);
        Assert.Contains("ET", page.ContentOperators);
        Assert.False(page.ContentOperatorsTruncated);
        Assert.Contains(report.Objects, item => item.DecodedStreamPreview?.Contains("BT", StringComparison.Ordinal) == true);
        Assert.Contains("PDF DEBUG DUMP", text, StringComparison.Ordinal);
        Assert.Contains("REV 2", text, StringComparison.Ordinal);
        Assert.Contains("PAGE 1", text, StringComparison.Ordinal);
        Assert.Contains("operators=[", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Dump_BoundsOperatorsAndReadsEncryptedDocumentsWithExplicitOptions() {
        byte[] encrypted = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted debugger"))
            .ToBytes();

        PdfDebuggerReport report = PdfDebugger.Dump(
            encrypted,
            new PdfDebuggerOptions { MaxContentOperatorsPerPage = 2 },
            new PdfReadOptions { Password = "open" });

        PdfDebugPage page = Assert.Single(report.Pages);
        Assert.Equal(2, page.ContentOperators.Count);
        Assert.True(page.ContentOperatorsTruncated);
        Assert.Throws<PdfPasswordRequiredException>(() => PdfDebugger.Dump(encrypted));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDebugger.Dump(encrypted, new PdfDebuggerOptions { MaxContentOperatorsPerPage = 0 }, new PdfReadOptions { Password = "open" }));
    }
}
