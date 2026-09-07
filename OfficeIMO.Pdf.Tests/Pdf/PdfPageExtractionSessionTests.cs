using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPageExtractionSessionTests {
    [Fact]
    public void ReadModelEditsDoNotChangeMetadataPreservedByPageOperations() {
        byte[] bytes = PdfDocument.Create().Meta(title: "Original title")
            .Paragraph(p => p.Text("FirstPage")).PageBreak().Paragraph(p => p.Text("SecondPage")).ToBytes();
        PdfDocument source = PdfDocument.Load(bytes);
        source.Read().Metadata.Title = "Changed read model";
        source.Inspect().Metadata.Author = "Changed inspection";

        var outputs = new List<PdfDocument> { source.Pages.Extract(1), PdfDocument.Merge(new[] { source, source }) };
        outputs.AddRange(source.Pages.Split());
        outputs.AddRange(source.Pages.Split(PdfPageSelection.From(2), PdfPageSelection.From(1)));
        foreach (PdfDocument output in outputs) {
            PdfMetadata metadata = PdfReadDocument.Open(output.ToBytes()).Metadata;
            Assert.Equal("Original title", metadata.Title);
            Assert.Null(metadata.Author);
        }
        Assert.Equal(bytes, source.ToBytes());
    }

    [Fact]
    public void CompoundSelectionsPreserveRepeatedPagesOrderAndSource() {
        byte[] bytes = PdfDocument.Create().Paragraph(p => p.Text("FirstPage"))
            .PageBreak().Paragraph(p => p.Text("SecondPage"))
            .PageBreak().Paragraph(p => p.Text("ThirdPage")).ToBytes();
        PdfDocument source = PdfDocument.Load(bytes);
        _ = source.Read();
        IReadOnlyList<PdfDocument> parts = source.Pages.Split(
            PdfPageSelection.From(3, 1), PdfPageSelection.From(2, 2), PdfPageSelection.From(1, 2, 3));
        string[][] expected = { new[] { "ThirdPage", "FirstPage" }, new[] { "SecondPage", "SecondPage" },
            new[] { "FirstPage", "SecondPage", "ThirdPage" } };
        Assert.Equal(expected.Length, parts.Count);
        for (int index = 0; index < parts.Count; index++) {
            using var independent = UglyToad.PdfPig.PdfDocument.Open(parts[index].ToBytes());
            Assert.Equal(expected[index], independent.GetPages().Select(static page => page.Text.Trim()).ToArray());
            Assert.Equal(expected[index].Length, parts[index].Inspect().PageCount);
        }
        Assert.Equal(bytes, source.ToBytes());
        Assert.Equal(3, source.Inspect().PageCount);
    }

    [Fact]
    public void WarmSourceStillEnforcesExplicitParserLimitsForPageOperations() {
        PdfDocument source = PdfDocument.Load(PdfDocument.Create().Paragraph(p => p.Text("Readable")).ToBytes());
        Assert.Contains("Readable", source.Reader.Text());
        var restricted = new PdfLoadOptions { Limits = new PdfReadLimits { MaxIndirectObjects = 1 } };
        Assert.Throws<PdfReadLimitException>(() => source.Pages.ExtractResult(PdfPageSelection.From(1), restricted));
        Assert.Throws<PdfReadLimitException>(() => source.Pages.SplitResult(options: restricted));
        Assert.Contains("Readable", source.Pages.Extract(1).Reader.Text());
    }

    [Fact]
    public void WarmEncryptedSourceDoesNotBypassAnOverridePassword() {
        byte[] bytes = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(p => p.Text("SecretPage")).ToBytes();
        PdfDocument source = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "owner" });
        Assert.Contains("SecretPage", source.Reader.Text());
        var incorrect = new PdfLoadOptions { Password = "incorrect" };
        Assert.Throws<PdfInvalidPasswordException>(() => source.Pages.ExtractResult(PdfPageSelection.From(1), incorrect));
        Assert.Throws<PdfInvalidPasswordException>(() => source.Pages.SplitResult(options: incorrect));
        Assert.Contains("SecretPage", source.Pages.Extract(1).Reader.Text());
    }
}
