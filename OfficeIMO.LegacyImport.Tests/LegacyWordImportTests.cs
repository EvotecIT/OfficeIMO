using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word.Legacy;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.OpenDocument;

namespace OfficeIMO.LegacyImport.Tests;

public sealed class LegacyWordImportTests {
    public static IEnumerable<object[]> Families() {
        yield return new object[] { LegacyFixtureFactory.WordPerfect(), "archive.wpd", LegacyWordFormat.WordPerfect };
        yield return new object[] { LegacyFixtureFactory.WordStar(), "archive.ws4", LegacyWordFormat.WordStar };
        yield return new object[] { LegacyFixtureFactory.AmiPro(), "archive.sam", LegacyWordFormat.AmiPro };
        yield return new object[] { LegacyFixtureFactory.WordPro(), "archive.lwp", LegacyWordFormat.LotusWordPro };
        yield return new object[] { LegacyFixtureFactory.WorksWord(), "archive.wps", LegacyWordFormat.MicrosoftWorks };
        yield return new object[] { LegacyFixtureFactory.Write(true), "archive.wri", LegacyWordFormat.MicrosoftWrite };
        yield return new object[] { LegacyFixtureFactory.Write(false), "archive.doc", LegacyWordFormat.WordForDos };
    }

    [Theory]
    [MemberData(nameof(Families))]
    public void DetectsAndImportsEveryBoundedFamily(byte[] source, string sourceName, LegacyWordFormat expected) {
        using LegacyWordImportResult result = LegacyWordImporter.Import(source, new LegacyWordImportOptions { SourceName = sourceName });
        Assert.Equal(expected, result.Detection.Format);
        Assert.NotEmpty(result.PlainText);
        Assert.True(result.Report.RecoveredItemCount > 0);
        Assert.NotNull(result.Document);
    }

    [Fact]
    public void ImportedWordModelUsesEverySupportedModernOutputOwner() {
        using LegacyWordImportResult imported = LegacyWordImporter.Import(LegacyFixtureFactory.WordStar(), new LegacyWordImportOptions { SourceName = "archive.ws4" });
        Assert.Equal(OfficeLegacyImportQuality.Structured, imported.Report.Quality);
        Assert.Single(imported.Document.Lists);
        using var docx = new MemoryStream();
        imported.Document.Save(docx);
        Assert.True(docx.Length > 100);
        Assert.Contains("First paragraph", imported.Document.ToHtml());
        Assert.Contains("First paragraph", imported.Document.ToMarkdown());
        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(imported.Document.ToPdf(), 0, 4));
        using var odt = new MemoryStream();
        imported.Document.ToOpenDocument().Save(odt);
        Assert.True(odt.Length > 100);
    }

    [Fact]
    public void LimitsAndStructuredPolicyFailClosed() {
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes("renamed plain text"),
            new LegacyWordImportOptions { SourceName = "renamed.wpd" }));
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(LegacyFixtureFactory.WordPerfect(), new LegacyWordImportOptions {
            SourceName = "archive.wpd", RequireStructured = true
        }));
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(new byte[9], new LegacyWordImportOptions {
            FormatHint = LegacyWordFormat.WordStar, Limits = new OfficeLegacyImportLimits { MaxInputBytes = 8 }
        }));

        using LegacyWordImportResult bounded = LegacyWordImporter.Import(LegacyFixtureFactory.WordStar(), new LegacyWordImportOptions {
            SourceName = "archive.ws4",
            Limits = new OfficeLegacyImportLimits { MaxInputBytes = int.MaxValue, MaxTextCharacters = 8 }
        });
        Assert.True(bounded.PlainText.Length <= 8);
    }

    [Fact]
    public void ImportHonorsCancellation() {
        Assert.Throws<OperationCanceledException>(() => LegacyWordImporter.Import(
            LegacyFixtureFactory.WordStar(),
            new LegacyWordImportOptions { SourceName = "archive.ws4" },
            new CancellationToken(canceled: true)));
    }

    [Fact]
    public void ReaderHandlerProjectsLegacyWarningsAndContent() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddLegacyWordHandler().Build();
        using var stream = new MemoryStream(LegacyFixtureFactory.WordPerfect());
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "archive.wpd");
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Recovered WordPerfect", StringComparison.Ordinal));
        Assert.Contains(OfficeDocumentReaderBuilderWordExtensions.LegacyHandlerId, result.CapabilitiesUsed);
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()), warning => warning.Contains("Legacy import quality", StringComparison.Ordinal));
    }

    [Fact]
    public void ReaderRegistrationCapturesLegacyImportOptions() {
        var options = new LegacyWordImportOptions {
            FormatHint = LegacyWordFormat.WordPerfect,
            Limits = new OfficeLegacyImportLimits { MaxInputBytes = 1024 }
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddLegacyWordHandler(options).Build();
        options.FormatHint = LegacyWordFormat.AmiPro;
        options.Limits.MaxInputBytes = 1;

        using var stream = new MemoryStream(LegacyFixtureFactory.WordPerfect());
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "archive.wpd");
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.Contains("wordperfect-5-6", StringComparison.Ordinal));
    }
}
