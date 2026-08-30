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
        Assert.Contains(imported.Content.Paragraphs.SelectMany(paragraph => paragraph.Runs), run => run.Bold && run.Text == "paragraph");
        Assert.Contains(imported.Content.Paragraphs, paragraph => paragraph.PageBreakBefore && paragraph.IsList);
        Assert.Contains(imported.Content.Notes, note => note.Kind == LegacyWordNoteKind.Comment && note.Text.Contains("Recovered comment", StringComparison.Ordinal));
        using var docx = new MemoryStream();
        imported.Document.Save(docx);
        Assert.True(docx.Length > 100);
        Assert.Contains("First", imported.Document.ToHtml());
        Assert.Contains("paragraph", imported.Document.ToHtml());
        Assert.Contains("First", imported.Document.ToMarkdown());
        Assert.Contains("paragraph", imported.Document.ToMarkdown());
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

        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(LegacyFixtureFactory.WordStar(), new LegacyWordImportOptions {
            SourceName = "archive.ws4",
            Limits = new OfficeLegacyImportLimits { MaxInputBytes = int.MaxValue, MaxTextCharacters = 8 }
        }));
    }

    [Fact]
    public void AmiProSam4RecoversStylesRunsAndParagraphLayout() {
        using LegacyWordImportResult imported = LegacyWordImporter.Import(LegacyFixtureFactory.AmiPro(), new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });
        Assert.Equal("4", imported.Metadata["AmiProVersion"]);
        Assert.Equal("1", imported.Metadata["StyleCount"]);
        LegacyWordParagraphContent first = imported.Content.Paragraphs[0];
        Assert.Equal("Body Text", first.StyleName);
        Assert.Contains(first.Runs, run => run.Bold && run.Text == "bold");
        Assert.True(first.KeepWithNext);
        Assert.Equal(12d, first.LineSpacingPoints);
        Assert.Contains(imported.Content.Paragraphs, paragraph => paragraph.Alignment == OfficeIMO.Word.WordParagraphAlignment.Center);
    }

    [Fact]
    public void WordStarDetectionRequiresCoherentGrammarAndHintedWeakInputIsSalvage() {
        byte[] arbitraryHighBit = Enumerable.Repeat((byte)0xC1, 128).ToArray();
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Detect(arbitraryHighBit, new LegacyWordImportOptions { SourceName = "random.ws4" }));

        using LegacyWordImportResult hinted = LegacyWordImporter.Import(arbitraryHighBit, new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar });
        Assert.Equal(OfficeLegacyImportQuality.Salvage, hinted.Report.Quality);
        Assert.Equal("wordstar-family-salvage", hinted.Report.SourceFormatId);
    }

    [Fact]
    public void WordStarPreservesExplicitEmptyParagraphsAndBoundsFormattedRuns() {
        byte[] paragraphs = Encoding.ASCII.GetBytes("\u0002\u0002One\r\n\r\nThree\r\n\u001A");
        using LegacyWordImportResult imported = LegacyWordImporter.Import(paragraphs, new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar, RequireStructured = true });
        Assert.Equal(3, imported.Content.Paragraphs.Count);
        Assert.Equal(string.Empty, imported.Content.Paragraphs[1].Text);

        var alternating = new List<byte>();
        for (int index = 0; index < 20; index++) { alternating.Add((byte)'A'); alternating.Add(0x02); }
        alternating.AddRange(new byte[] { 0x0D, 0x0A, 0x1A });
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(alternating.ToArray(), new LegacyWordImportOptions {
            FormatHint = LegacyWordFormat.WordStar,
            Limits = new OfficeLegacyImportLimits { MaxItems = 8 }
        }));

        var large = new List<byte>(300_010) { 0x02, 0x02 };
        large.AddRange(Enumerable.Repeat((byte)'A', 300_000));
        large.AddRange(new byte[] { 0x0D, 0x0A, 0x1A });
        using LegacyWordImportResult largeImported = LegacyWordImporter.Import(large.ToArray(), new LegacyWordImportOptions { SourceName = "large.ws7", RequireStructured = true });
        Assert.Equal(OfficeLegacyImportQuality.Structured, largeImported.Report.Quality);
    }

    [Fact]
    public void AmiProUnsupportedVersionsSalvageAndAllObjectSectionsAreInventoried() {
        byte[] version3 = Encoding.ASCII.GetBytes("[ver]\n3\n[edoc]\nLegacy Ami text\n");
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Detect(version3, new LegacyWordImportOptions { SourceName = "archive.sam" }));
        using LegacyWordImportResult salvage = LegacyWordImporter.Import(version3, new LegacyWordImportOptions { FormatHint = LegacyWordFormat.AmiPro });
        Assert.Equal(OfficeLegacyImportQuality.Salvage, salvage.Report.Quality);
        Assert.Equal("ami-pro-sam-unsupported-salvage", salvage.Report.SourceFormatId);

        byte[] sections = Encoding.ASCII.GetBytes("[ver]\n4\n[frm]\nframe payload\n[edoc]\nText\n[objdata]\nobject payload\n[lay]\nlayout payload\n");
        using LegacyWordImportResult structured = LegacyWordImporter.Import(sections, new LegacyWordImportOptions { SourceName = "archive.sam" });
        Assert.True(structured.Report.InertContent.HasFlag(OfficeLegacyInertContentKind.EmbeddedObjects));
        Assert.Contains(structured.Report.Findings, finding => finding.Code == "AMIPRO_EMBEDDED_OBJECT_INERT");
        Assert.Contains(structured.Report.Findings, finding => finding.Code == "AMIPRO_SECTION_UNSUPPORTED");
    }

    [Fact]
    public void AmiProFormattedRunsAndUnknownTagKindsRespectItemLimit() {
        var source = new StringBuilder("[ver]\n4\n[edoc]\n");
        for (int index = 0; index < 20; index++) source.Append("A<+!>B<-!>");
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(Encoding.ASCII.GetBytes(source.ToString()), new LegacyWordImportOptions {
            SourceName = "archive.sam",
            Limits = new OfficeLegacyImportLimits { MaxItems = 8 }
        }));

        source.Clear().Append("[ver]\n4\n[edoc]\nText");
        for (int index = 0; index < 20; index++) source.Append('<').Append("unknown").Append(index).Append('>');
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(Encoding.ASCII.GetBytes(source.ToString()), new LegacyWordImportOptions {
            SourceName = "archive.sam",
            Limits = new OfficeLegacyImportLimits { MaxItems = 8 }
        }));

        string overlongSection = "[" + new string('A', 16_384) + "]";
        using LegacyWordImportResult boundedDiagnostic = LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\nText\n" + overlongSection + "\npayload\n"),
            new LegacyWordImportOptions { SourceName = "archive.sam", Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 16 } });
        Assert.Contains(boundedDiagnostic.Report.Findings, finding => finding.Code == "AMIPRO_SECTION_UNSUPPORTED" && finding.Message.Length < 256);
    }

    [Fact]
    public void WordStarExternalGraphicsStayInertAndMalformedSequencesFailClosed() {
        using LegacyWordImportResult imported = LegacyWordImporter.Import(LegacyFixtureFactory.WordStarWithGraphics(), new LegacyWordImportOptions { SourceName = "archive.ws7", FormatHint = LegacyWordFormat.WordStar });
        LegacyWordResourceReference resource = Assert.Single(imported.Content.Resources);
        Assert.Equal("Graphics", resource.Kind);
        Assert.EndsWith("FIGURE.PCX", resource.Reference, StringComparison.Ordinal);
        Assert.True(imported.Report.InertContent.HasFlag(OfficeLegacyInertContentKind.ExternalLinks));

        byte[] malformed = { (byte)'T', 0x1D, 0x20, 0, 0x06, (byte)'X', 0x1A };
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(malformed, new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar }));
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
