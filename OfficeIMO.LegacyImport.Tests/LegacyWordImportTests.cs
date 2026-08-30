using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using OfficeIMO.Word.Legacy;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.OpenDocument;
using System.Threading.Tasks;

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
    public void StructuredWordStarParagraphsEnforceTheRecordLimit() {
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes("One\r\nTwo\r\n\x1A"),
            new LegacyWordImportOptions {
                FormatHint = LegacyWordFormat.WordStar,
                RequireStructured = true,
                Limits = new OfficeLegacyImportLimits { MaxRecords = 1, MaxItems = 100 }
            }));
    }

    [Fact]
    public void AmiProSam4RecoversStylesRunsAndParagraphLayout() {
        using LegacyWordImportResult imported = LegacyWordImporter.Import(LegacyFixtureFactory.AmiPro(), new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });
        Assert.Equal("4", imported.Metadata["AmiProVersion"]);
        Assert.Equal("1", imported.Metadata["StyleCount"]);
        LegacyWordStyleContent style = Assert.Single(imported.Content.Styles);
        Assert.Equal("Body Text", style.Name);
        Assert.Equal("Arial", style.FontFamily);
        Assert.Equal(12d, style.FontSizePoints);
        Assert.Equal("FF0000", style.ColorHex);
        Assert.True(style.Bold);
        Assert.Equal(OfficeIMO.Word.WordParagraphAlignment.Left, style.Alignment);
        Assert.Equal(12d, style.LineSpacingPoints);
        Assert.True(style.KeepWithNext);
        Assert.True(style.KeepLinesTogether);
        LegacyWordParagraphContent first = imported.Content.Paragraphs[0];
        Assert.Equal("Body Text", first.StyleName);
        Assert.Contains(first.Runs, run => run.Bold && run.Text == "bold");
        Assert.True(first.KeepWithNext);
        Assert.Equal(12d, first.LineSpacingPoints);
        Assert.Contains(imported.Content.Paragraphs, paragraph => paragraph.Alignment == OfficeIMO.Word.WordParagraphAlignment.Center);

        DocumentFormat.OpenXml.Wordprocessing.Style projectedStyle = imported.Document.OpenXmlDocument.MainDocumentPart!
            .StyleDefinitionsPart!.Styles!.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>()
            .Single(candidate => candidate.StyleName?.Val?.Value == "Body Text");
        Assert.Equal("Arial", projectedStyle.StyleRunProperties!.RunFonts!.Ascii!.Value);
        Assert.Equal("24", projectedStyle.StyleRunProperties.FontSize!.Val!.Value);
        Assert.Equal("FF0000", projectedStyle.StyleRunProperties.Color!.Val!.Value);
        Assert.True(projectedStyle.StyleRunProperties.Bold!.Val!.Value);
        Assert.False(projectedStyle.StyleRunProperties.Italic!.Val!.Value);
        Assert.Equal(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Left,
            projectedStyle.StyleParagraphProperties!.Justification!.Val!.Value);
        Assert.Equal("240", projectedStyle.StyleParagraphProperties.SpacingBetweenLines!.Line!.Value);
        Assert.Equal("0", projectedStyle.StyleParagraphProperties.SpacingBetweenLines.Before!.Value);
        Assert.Equal("0", projectedStyle.StyleParagraphProperties.SpacingBetweenLines.After!.Value);
        Assert.NotNull(projectedStyle.StyleParagraphProperties.KeepNext);
        Assert.NotNull(projectedStyle.StyleParagraphProperties.KeepLines);
    }

    [Fact]
    public void EmptyAmiProSam4ProjectsAnEmptyDocument() {
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\n"),
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

        Assert.Equal(OfficeLegacyImportQuality.Structured, imported.Report.Quality);
        Assert.Empty(imported.Content.Paragraphs);
        Assert.Equal(string.Empty, imported.PlainText);
        imported.Report.RequireStructuredNoLoss();
        Assert.Single(imported.Document.Paragraphs);
    }

    [Fact]
    public void AmiProRetainsFormattingOnlyParagraphs() {
        foreach (string paragraphSource in new[] { "<+B>", "@Body Text@" }) {
            string source = Encoding.ASCII.GetString(LegacyFixtureFactory.AmiPro());
            source = source.Substring(0, source.IndexOf("[edoc]", StringComparison.Ordinal)) +
                "[edoc]\n" + paragraphSource + "\n";

            using LegacyWordImportResult imported = LegacyWordImporter.Import(
                Encoding.ASCII.GetBytes(source),
                new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

            LegacyWordParagraphContent paragraph = Assert.Single(imported.Content.Paragraphs);
            Assert.Equal(string.Empty, paragraph.Text);
            Assert.True(paragraph.Alignment == WordParagraphAlignment.Center || paragraph.StyleName == "Body Text");
            Assert.Single(imported.Document.Paragraphs);
            imported.Report.RequireStructuredNoLoss();
        }
    }

    [Fact]
    public void AmiProRejectsUnsafeDecodedEscapeCharacters() {
        foreach (string escaped in new[] { "<\\\t>", "</@>" }) {
            using LegacyWordImportResult imported = LegacyWordImporter.Import(
                Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\n" + escaped + "Visible\n"),
                new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

            Assert.Equal("Visible", imported.PlainText);
            Assert.Contains(imported.Report.Findings, finding => finding.Code == "AMIPRO_INLINE_TAG_MALFORMED");
            Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
            using var docx = new MemoryStream();
            imported.Document.Save(docx);
            Assert.True(docx.Length > 0);
        }
    }

    [Fact]
    public void AmiProProjectionWritesExplicitOverridesForStyleResets() {
        string source = Encoding.ASCII.GetString(LegacyFixtureFactory.AmiPro())
            .Replace("\n16385\n", "\n16391\n", StringComparison.Ordinal)
            .Replace("<+!>bold<-!> paragraph", "<+!>bold<-!><-\"><-#><:f> reset", StringComparison.Ordinal)
            .Replace("\n\n<+B>- Ami list\n", "\n", StringComparison.Ordinal);

        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes(source),
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

        WordDocument document = Assert.IsType<WordDocument>(imported.Document);
        DocumentFormat.OpenXml.Wordprocessing.Run resetRun = document.OpenXmlDocument!
            .MainDocumentPart!.Document!.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
            .Single(run => run.InnerText == " reset");
        DocumentFormat.OpenXml.Wordprocessing.RunProperties properties = resetRun.RunProperties!;
        Assert.False(properties.Bold!.Val!.Value);
        Assert.False(properties.Italic!.Val!.Value);
        Assert.Equal(DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.None, properties.Underline!.Val!.Value);
        Assert.Equal(DocumentFormat.OpenXml.Wordprocessing.ThemeFontValues.MinorHighAnsi, properties.RunFonts!.AsciiTheme!.Value);
        Assert.Equal("22", properties.FontSize!.Val!.Value);
        Assert.Equal("auto", properties.Color!.Val!.Value);
        imported.Report.RequireStructuredNoLoss();
    }

    [Fact]
    public void AmiProDocumentDirectivesRemainExplicitLoss() {
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\n>directive payload\nVisible\n"),
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

        Assert.Equal("Visible", imported.PlainText);
        Assert.Equal("1", imported.Metadata["AmiProDocumentDirectiveCount"]);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "AMIPRO_DOCUMENT_DIRECTIVE_UNSUPPORTED");
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
    }

    [Fact]
    public void AmiProUnsupportedStyleFlagsRemainExplicitLoss() {
        string source = Encoding.ASCII.GetString(LegacyFixtureFactory.AmiPro())
            .Replace("\n16385\n", "\n16513\n", StringComparison.Ordinal)
            .Replace("[algn]\n1\n", "[algn]\n17\n", StringComparison.Ordinal)
            .Replace("[spc]\n1\n", "[spc]\n17\n", StringComparison.Ordinal)
            .Replace("[brk]\n16\n", "[brk]\n18\n", StringComparison.Ordinal);
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes(source),
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

        Assert.Equal("0x80", imported.Metadata["AmiProUnsupportedStyleFlags.Formatting"]);
        Assert.Equal("0x10", imported.Metadata["AmiProUnsupportedStyleFlags.Alignment"]);
        Assert.Equal("0x10", imported.Metadata["AmiProUnsupportedStyleFlags.Spacing"]);
        Assert.Equal("0x2", imported.Metadata["AmiProUnsupportedStyleFlags.Break"]);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "AMIPRO_STYLE_FLAGS_UNSUPPORTED");
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
    }

    [Fact]
    public void WordStarPreservesRepeatedAndTrailingPageBreaks() {
        byte[] source = new byte[] { 0x02, 0x02 }
            .Concat(Encoding.ASCII.GetBytes("First\r\n"))
            .Concat(new byte[] { 0x0C, 0x0C })
            .Concat(Encoding.ASCII.GetBytes("Second\r\n"))
            .Concat(new byte[] { 0x0C, 0x1A })
            .ToArray();
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar, RequireStructured = true });

        Assert.Equal(3, imported.Content.Paragraphs.Count(paragraph => paragraph.PageBreakBefore));
        Assert.True(imported.Content.Paragraphs[^1].PageBreakBefore);
        imported.Report.RequireStructuredNoLoss();
    }

    [Fact]
    public void AmiProStyleMetadataAndInlineReferencesShareTheTextBudget() {
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            LegacyFixtureFactory.AmiPro(),
            new LegacyWordImportOptions {
                SourceName = "archive.sam",
                Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 13 }
            }));

        byte[] inlineReference = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\n@S@X\n");
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            inlineReference,
            new LegacyWordImportOptions {
                SourceName = "archive.sam",
                Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 1 }
            }));
    }

    [Fact]
    public void AmiProMalformedStyleBlocksAndInlineValuesAreReportedAsLoss() {
        byte[] source = Encoding.ASCII.GetBytes("[ver]\n4\n[tag]\nBroken\n[edoc]\nText<:S+bad><:fbad>\n");
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

        Assert.Contains(imported.Report.Findings, finding => finding.Code == "AMIPRO_STYLE_BLOCK_MALFORMED");
        Assert.Single(imported.Report.Findings, finding => finding.Code == "AMIPRO_INLINE_TAG_MALFORMED");
        Assert.Equal("1", imported.Metadata["AmiProMalformedStyleBlockCount"]);
        Assert.Equal("2", imported.Metadata["AmiProMalformedInlineTagCount"]);
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
    }

    [Theory]
    [InlineData("<:f0,Arial,0,0,0>")]
    [InlineData("<:f-20,Arial,0,0,0>")]
    [InlineData("<:f240,Arial,999,0,0>")]
    [InlineData("<:S+0>")]
    [InlineData("<:S+-4>")]
    public void AmiProInlineMeasurementsRejectInvalidValuesWithoutPartialFormatting(string tag) {
        byte[] source = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\n" + tag + "X\n");
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

        LegacyWordParagraphContent paragraph = Assert.Single(imported.Content.Paragraphs);
        LegacyWordRunContent run = Assert.Single(paragraph.Runs);
        Assert.Null(run.FontSizePoints);
        Assert.Null(run.FontFamily);
        Assert.Null(run.ColorHex);
        Assert.Null(paragraph.LineSpacingPoints);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "AMIPRO_INLINE_TAG_MALFORMED");
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
    }

    [Theory]
    [InlineData("Body Text\n0\n[fnt]", "Body Text\nbad\n[fnt]")]
    [InlineData("240\n255\n16385", "240\nbad\n16385")]
    [InlineData("255\n16385\n[algn]", "255\nbad\n[algn]")]
    [InlineData("[algn]\n1\n0\n0\n0\n0", "[algn]\nbad\n0\n0\n0\n0")]
    [InlineData("[algn]\n1\n0\n0\n0\n0", "[algn]\n1\nbad\n0\n0\n0")]
    [InlineData("[spc]\n1\n240\n0\n0\n0", "[spc]\nbad\n240\n0\n0\n0")]
    [InlineData("[spc]\n1\n240\n0\n0\n0", "[spc]\n1\nbad\n0\n0\n0")]
    [InlineData("[spc]\n1\n240\n0\n0\n0", "[spc]\n1\n240\nbad\n0\n0")]
    [InlineData("[spc]\n1\n240\n0\n0\n0", "[spc]\n1\n240\n0\nbad\n0")]
    [InlineData("[spc]\n1\n240\n0\n0\n0", "[spc]\n1\n240\n0\n0\nbad")]
    [InlineData("[spc]\n1\n240\n0\n0\n0", "[spc]\n1\n240\n0\n-1\n0")]
    [InlineData("[spc]\n1\n240\n0\n0\n0", "[spc]\n1\n240\n0\n0\n-1")]
    [InlineData("[brk]\n16\n[edoc]", "[brk]\nbad\n[edoc]")]
    public void AmiProMalformedStyleNumericFieldsRemainExplicitLoss(string sourceValue, string replacement) {
        string fixture = Encoding.ASCII.GetString(LegacyFixtureFactory.AmiPro());
        string malformed = fixture.Replace(sourceValue, replacement, StringComparison.Ordinal);
        Assert.NotEqual(fixture, malformed);

        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes(malformed),
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

        Assert.Equal("1", imported.Metadata["AmiProMalformedStyleBlockCount"]);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "AMIPRO_STYLE_BLOCK_MALFORMED");
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
    }

    [Fact]
    public void AmiProStyleBlocksRejectUnconsumedFields() {
        string fixture = Encoding.ASCII.GetString(LegacyFixtureFactory.AmiPro());
        string malformed = fixture.Replace("[brk]\n16\n[edoc]", "[brk]\n16\nunconsumed\n[edoc]", StringComparison.Ordinal);

        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes(malformed),
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

        Assert.Empty(imported.Content.Styles);
        Assert.Equal("1", imported.Metadata["AmiProMalformedStyleBlockCount"]);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "AMIPRO_STYLE_BLOCK_MALFORMED");
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
    }

    [Theory]
    [InlineData("[ver]\n4\n[ver]\n4\n[edoc]\nOne\n")]
    [InlineData("[ver]\n4\n[edoc]\nOne\n[edoc]\nTwo\n")]
    public void AmiProRejectsDuplicateSingletonSections(string source) {
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes(source),
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true }));
    }

    [Fact]
    public void AmiProUnterminatedInlineOpenersAreScannedLinearlyWithoutHidingLaterStyleReferences() {
        string text = string.Concat(Enumerable.Repeat("<x", 100_000));
        byte[] source = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\n" + text + "@Missing@End\n");

        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });

        Assert.Equal(text + "End", imported.PlainText);
    }

    [Fact]
    public void AmiProDuplicateStyleDefinitionsCannotDisappearFromNoLossClaims() {
        string fixture = Encoding.ASCII.GetString(LegacyFixtureFactory.AmiPro());
        int styleStart = fixture.IndexOf("[tag]", StringComparison.Ordinal);
        int documentStart = fixture.IndexOf("[edoc]", StringComparison.Ordinal);
        string duplicate = fixture.Substring(styleStart, documentStart - styleStart);
        byte[] source = Encoding.ASCII.GetBytes(fixture.Insert(documentStart, duplicate));

        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });
        Assert.Equal("1", imported.Metadata["AmiProDuplicateStyleCount"]);
        Assert.Single(imported.Report.Findings, finding => finding.Code == "AMIPRO_STYLE_DUPLICATE");
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
    }

    [Fact]
    public void AmiProPreservesHalfPointStyleAndInlineFontSizes() {
        string styledSource = Encoding.ASCII.GetString(LegacyFixtureFactory.AmiPro())
            .Replace("[fnt]\nArial\n240\n", "[fnt]\nArial\n210\n", StringComparison.Ordinal);
        using LegacyWordImportResult styled = LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes(styledSource),
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });
        Assert.Equal(10.5d, styled.Content.Paragraphs[0].Runs[0].FontSizePoints);

        byte[] inlineSource = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\n<:f210,Arial,0,0,0>X\n");
        using LegacyWordImportResult inline = LegacyWordImporter.Import(
            inlineSource,
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });
        Assert.Equal(10.5d, inline.Content.Paragraphs[0].Runs[0].FontSizePoints);
    }

    [Fact]
    public void AmiProStructuredProfileRejectsUndeclaredExtendedTextEncoding() {
        byte[] source = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\nX\n");
        source[Array.IndexOf(source, (byte)'X')] = 0xE9;
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { SourceName = "archive.sam" }));
    }

    [Theory]
    [InlineData((byte)0x00)]
    [InlineData((byte)0x01)]
    [InlineData((byte)0x0B)]
    [InlineData((byte)0x0C)]
    public void AmiProStructuredProfileRejectsXmlInvalidAsciiControls(byte control) {
        byte[] source = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\nXY\n");
        source[Array.IndexOf(source, (byte)'Y')] = control;

        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true }));
    }

    [Fact]
    public void AmiProLargeInlineTagScansObserveCancellation() {
        string tags = string.Concat(Enumerable.Repeat("<+!><-!>", 2_000_000));
        byte[] source = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\n" + tags + "X\n");
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(1));

        Assert.Throws<OperationCanceledException>(() => LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true },
            cancellation.Token));
    }

    [Fact]
    public void CompoundDetectionHonorsRaisedDirectoryEntryLimit() {
        byte[] source = LegacyFixtureFactory.CompoundWithLargeDirectory();
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Detect(
            source,
            new LegacyWordImportOptions { SourceName = "archive.lwp" }));

        LegacyWordDetection detected = LegacyWordImporter.Detect(
            source,
            new LegacyWordImportOptions {
                SourceName = "archive.lwp",
                Limits = new OfficeLegacyImportLimits { MaxCompoundStreams = 17 * 32 }
            });

        Assert.Equal(LegacyWordFormat.LotusWordPro, detected.Format);
    }

    [Fact]
    public void WordStarDetectionRequiresCoherentGrammarAndHintedWeakInputIsSalvage() {
        byte[] arbitraryHighBit = Enumerable.Repeat((byte)0xC1, 128).ToArray();
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Detect(arbitraryHighBit, new LegacyWordImportOptions { SourceName = "random.ws4" }));

        using LegacyWordImportResult hinted = LegacyWordImporter.Import(arbitraryHighBit, new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar });
        Assert.Equal(OfficeLegacyImportQuality.Salvage, hinted.Report.Quality);
        Assert.Equal("wordstar-family-salvage", hinted.Report.SourceFormatId);
    }

    [Theory]
    [InlineData((byte)0x0D)]
    [InlineData((byte)0x0A)]
    public void WordStarAdmissionAcceptsEveryStructuredHardParagraphTerminator(byte terminator) {
        byte[] source = new byte[] { 0x02, 0x02 }
            .Concat(Encoding.ASCII.GetBytes("One"))
            .Append(terminator)
            .Concat(Encoding.ASCII.GetBytes("Two"))
            .Append(terminator)
            .Append((byte)0x1A)
            .ToArray();

        LegacyWordDetection detected = LegacyWordImporter.Detect(source);
        Assert.Equal(LegacyWordFormat.WordStar, detected.Format);
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { RequireStructured = true });
        Assert.Equal(new[] { "One", "Two" }, imported.Content.Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public void WorksShortHeaderRequiresCorroboratingSourceEvidence() {
        byte[] source = LegacyFixtureFactory.WorksWord();
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Detect(source));
        Assert.Equal(LegacyWordFormat.MicrosoftWorks,
            LegacyWordImporter.Detect(source, new LegacyWordImportOptions { SourceName = "archive.wps" }).Format);
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
    public void WordStarParagraphStyleSequencesAttachToFollowingParagraphs() {
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            LegacyFixtureFactory.WordStarWithStyles(),
            new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar, RequireStructured = true });

        Assert.Collection(imported.Content.Paragraphs,
            paragraph => {
                Assert.Equal("Body paragraph", paragraph.Text);
                Assert.Equal("Body", paragraph.StyleName);
            },
            paragraph => {
                Assert.Equal("Heading paragraph", paragraph.Text);
                Assert.Equal("Heading", paragraph.StyleName);
            });

        WordParagraphSnapshot[] projected = imported.Document.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements)
            .OfType<WordParagraphSnapshot>()
            .ToArray();
        Assert.Collection(projected,
            paragraph => {
                Assert.Equal("Body paragraph", paragraph.Text);
                Assert.Equal("Body", paragraph.StyleName);
            },
            paragraph => {
                Assert.Equal("Heading paragraph", paragraph.Text);
                Assert.Equal("Heading", paragraph.StyleName);
            });
    }

    [Fact]
    public void RecoveredStylesDoNotReuseApplicationRegisteredStyleIds() {
        const string sourceStyleName = "BodyCollisionProbe2405";
        const string registeredStyleId = "LegacyBodyCollisionProbe2405";
        WordParagraphStyle.RegisterCustomStyle(registeredStyleId, new WordParagraphStyleDefinition(registeredStyleId) {
            Name = "Unrelated application style",
            Bold = true
        });

        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            LegacyFixtureFactory.WordStarWithStyle(sourceStyleName),
            new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar, RequireStructured = true });

        DocumentFormat.OpenXml.Packaging.WordprocessingDocument package = imported.Document.OpenXmlDocument
            ?? throw new InvalidDataException("Imported document has no Open XML package.");
        DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart = package.MainDocumentPart
            ?? throw new InvalidDataException("Imported document has no main document part.");
        DocumentFormat.OpenXml.Wordprocessing.Document mainDocument = mainPart.Document
            ?? throw new InvalidDataException("Imported document has no document root.");
        DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = mainDocument.Body!
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Single();
        Assert.Equal(registeredStyleId + "2", paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        DocumentFormat.OpenXml.Wordprocessing.Style projectedStyle = mainPart.StyleDefinitionsPart!.Styles!
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Style>()
            .Single(style => style.StyleId?.Value == registeredStyleId + "2");
        Assert.Equal(sourceStyleName, projectedStyle.StyleName!.Val!.Value);
    }

    [Fact]
    public void WordStarStyleSequenceTextCannotExceedTheSharedCharacterBudget() {
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            LegacyFixtureFactory.WordStarWithStyle(new string('S', 32)),
            new LegacyWordImportOptions {
                FormatHint = LegacyWordFormat.WordStar,
                Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 16 }
            }));
    }

    [Fact]
    public void WordStarCoalescesUnknownDotCommandsAndPlainTextRemainsHardBounded() {
        byte[] commands = Encoding.ASCII.GetBytes("\u0002\u0002.XX first\r\n.YY second\r\nText\r\n\u001A");
        using LegacyWordImportResult imported = LegacyWordImporter.Import(commands, new LegacyWordImportOptions {
            FormatHint = LegacyWordFormat.WordStar,
            RequireStructured = true
        });
        Assert.Single(imported.Report.Findings, finding => finding.Code == "WORDSTAR_DOT_COMMAND");
        Assert.Equal("2", imported.Metadata["WordStarUnknownDotCommandCount"]);

        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes("A\r\nB\r\n\u001A"),
            new LegacyWordImportOptions {
                FormatHint = LegacyWordFormat.WordStar,
                Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 2 }
            }));
    }

    [Fact]
    public void WordStarUnsupportedControlsAndMetadataOnlyHeadersAreReportedAsLoss() {
        byte[] source = Encoding.ASCII.GetBytes("\u0002\u0002.HE Header\r\n.FO Footer\r\nBody").Concat(new byte[] { 0x07, 0x0D, 0x0A, 0x1A }).ToArray();
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar, RequireStructured = true });

        Assert.Equal("Header", imported.Metadata["Header"]);
        Assert.Equal("Footer", imported.Metadata["Footer"]);
        Assert.Equal("2", imported.Metadata["WordStarHeaderFooterMetadataOnlyCount"]);
        Assert.Equal("1", imported.Metadata["WordStarUnsupportedControl.0x07Count"]);
        Assert.Single(imported.Report.Findings, finding => finding.Code == "WORDSTAR_HEADER_FOOTER_METADATA_ONLY");
        Assert.Single(imported.Report.Findings, finding => finding.Code == "WORDSTAR_CONTROL_UNSUPPORTED");
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
    }

    [Fact]
    public void WordStarRejectsNonPaddingDataAfterEof() {
        byte[] source = Encoding.ASCII.GetBytes("\u0002\u0002Text\r\n\u001ATrailing");
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            source,
            new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar }));
    }

    [Fact]
    public void WordStarCoalescesRepeatedSequenceResourceAndListDiagnostics() {
        using LegacyWordImportResult imported = LegacyWordImporter.Import(
            LegacyFixtureFactory.WordStarWithRepeatedDiagnostics(),
            new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar, RequireStructured = true });

        Assert.Single(imported.Report.Findings, finding => finding.Code == "WORDSTAR_SEQUENCE_PARTIAL");
        Assert.Single(imported.Report.Findings, finding => finding.Code == "WORDSTAR_SEQUENCE_UNSUPPORTED");
        Assert.Single(imported.Report.Findings, finding => finding.Code == "WORDSTAR_GRAPHICS_REFERENCE_INERT");
        Assert.Single(imported.Report.Findings, finding => finding.Code == "WORDSTAR_LIST_INFERRED");
        Assert.Equal("3", imported.Metadata["WordStarPartialSequence.0x00Count"]);
        Assert.Equal("3", imported.Metadata["WordStarUnsupportedSequence.0x20Count"]);
        Assert.Equal("3", imported.Metadata["WordStarGraphicsReferenceCount"]);
        Assert.Equal("2", imported.Metadata["WordStarInferredListCount"]);
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
    public void AmiProIndentedSectionBoundariesRemainInert() {
        byte[] source = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\nVisible text\n  [objdata]  \nobject payload\n");
        using LegacyWordImportResult imported = LegacyWordImporter.Import(source, new LegacyWordImportOptions {
            SourceName = "archive.sam",
            RequireStructured = true
        });

        Assert.Equal("Visible text", imported.PlainText);
        Assert.DoesNotContain("objdata", imported.PlainText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("object payload", imported.PlainText, StringComparison.OrdinalIgnoreCase);
        Assert.True(imported.Report.InertContent.HasFlag(OfficeLegacyInertContentKind.EmbeddedObjects));
    }

    [Fact]
    public void SalvageWordParsingEnforcesTheRecordLimit() {
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            LegacyFixtureFactory.WordPerfect(),
            new LegacyWordImportOptions {
                FormatHint = LegacyWordFormat.WordPerfect,
                Limits = new OfficeLegacyImportLimits { MaxRecords = 1, MaxItems = 100 }
            }));
    }

    [Fact]
    public void AmiProVersionMarkerMustStartTheHeader() {
        byte[] prefixed = Encoding.ASCII.GetBytes("unrelated data\n[ver]\n4\n[edoc]\nText\n");
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Detect(
            prefixed,
            new LegacyWordImportOptions { SourceName = "archive.sam" }));

        using LegacyWordImportResult whitespace = LegacyWordImporter.Import(
            Encoding.ASCII.GetBytes(" \t\n[ver]\n4\n[edoc]\nText\n"),
            new LegacyWordImportOptions { SourceName = "archive.sam", RequireStructured = true });
        Assert.Equal(OfficeLegacyImportQuality.Structured, whitespace.Report.Quality);
    }

    [Fact]
    public void AmiProRecordLimitIsCheckedBeforeLineMaterialization() {
        byte[] source = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\nOne\nTwo\nThree\n");
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(source, new LegacyWordImportOptions {
            SourceName = "archive.sam",
            Limits = new OfficeLegacyImportLimits { MaxRecords = 4 }
        }));
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

    [Theory]
    [InlineData((byte)0x03)]
    [InlineData((byte)0x04)]
    [InlineData((byte)0x05)]
    [InlineData((byte)0x06)]
    [InlineData((byte)0x10)]
    [InlineData((byte)0x11)]
    public void WordStarTextBearingSequencesRejectUnsupportedControls(byte sequenceType) {
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            LegacyFixtureFactory.WordStarSequenceWithControl(sequenceType, 0x07),
            new LegacyWordImportOptions { FormatHint = LegacyWordFormat.WordStar, RequireStructured = true }));
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
    public void ReaderHandlerPreservesLegacyWarningsWhenProjectionHasNoChunks() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddLegacyWordHandler().Build();
        byte[] source = Encoding.ASCII.GetBytes("[ver]\n4\n[edoc]\n>unsupported directive\n");

        OfficeDocumentReadResult result = reader.ReadDocument(source, "archive.sam");

        Assert.Empty(result.Chunks);
        Assert.Contains(OfficeDocumentReaderBuilderWordExtensions.LegacyHandlerId, result.CapabilitiesUsed);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Message.Contains("Legacy import quality", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Message.Contains("AMIPRO_DOCUMENT_DIRECTIVE_UNSUPPORTED", StringComparison.Ordinal));
    }

    [Fact]
    public void AmiProRemainsOnLegacyWordHandlerInPreferContentMode() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddWordAndLegacyHandlers()
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            LegacyFixtureFactory.AmiPro(),
            "archive.sam",
            new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });

        Assert.Contains(OfficeDocumentReaderBuilderWordExtensions.LegacyHandlerId, result.CapabilitiesUsed);
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Ami Pro", StringComparison.Ordinal));
    }

    [Fact]
    public void WordStarRemainsOnLegacyWordHandlerInPreferContentMode() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddWordAndLegacyHandlers()
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            LegacyFixtureFactory.WordStarMarkdownLike(),
            "archive.ws4",
            new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });

        Assert.Contains(OfficeDocumentReaderBuilderWordExtensions.LegacyHandlerId, result.CapabilitiesUsed);
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("WordStar heading", StringComparison.Ordinal));
    }

    [Fact]
    public void WordReaderContentRoutesStrongDosDocHeadersWithoutStealingCompoundDoc() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddWordAndLegacyHandlers()
            .Build();
        using var stream = new MemoryStream(LegacyFixtureFactory.Write(write: false));

        OfficeDocumentReadResult result = reader.ReadDocument(stream, "archive.doc");

        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Word DOS recovered paragraph", StringComparison.Ordinal));
        Assert.Contains(OfficeDocumentReaderBuilderWordExtensions.LegacyHandlerId, result.CapabilitiesUsed);
    }

    [Fact]
    public void NormalWordHandlerDoesNotOptIntoWordForDosSalvage() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

        Assert.ThrowsAny<Exception>(() => reader.ReadDocument(
            LegacyFixtureFactory.Write(false),
            "archive.doc"));
    }

    [Fact]
    public void CoordinatedWordRegistrationAppliesLegacyOptionsToDosDocRouting() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddWordAndLegacyHandlers(new LegacyWordImportOptions {
                RequireStructured = true
            })
            .Build();
        using var stream = new MemoryStream(LegacyFixtureFactory.Write(write: false));

        Assert.Throws<InvalidDataException>(() => reader.ReadDocument(stream, "archive.doc"));
    }

    [Fact]
    public void CoordinatedWordRegistrationBoundsNonSeekableDosRoutingBeforeSnapshot() {
        byte[] source = LegacyFixtureFactory.Write(write: false);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddWordAndLegacyHandlers(new LegacyWordImportOptions {
                Limits = new OfficeLegacyImportLimits { MaxInputBytes = 128 }
            })
            .Build();
        using var stream = new NonSeekableStream(source);

        IOException exception = Assert.Throws<IOException>(() => reader.ReadDocument(stream, "archive.doc"));
        Assert.Contains("MaxInputBytes", exception.Message, StringComparison.Ordinal);
        Assert.Equal(97, stream.RequestedCounts[0]);
        Assert.All(stream.RequestedCounts.Skip(1), count => Assert.InRange(count, 1, 32));
        Assert.Equal(129, stream.BytesRead);
    }

    [Fact]
    public async Task CoordinatedWordRegistrationBoundsNonSeekableDosRoutingBeforeAsyncSnapshot() {
        byte[] source = LegacyFixtureFactory.Write(write: false);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddWordAndLegacyHandlers(new LegacyWordImportOptions {
                Limits = new OfficeLegacyImportLimits { MaxInputBytes = 128 }
            })
            .Build();
        using var stream = new NonSeekableStream(source);

        IOException exception = await Assert.ThrowsAsync<IOException>(() => reader.ReadDocumentAsync(stream, "archive.doc"));
        Assert.Contains("MaxInputBytes", exception.Message, StringComparison.Ordinal);
        Assert.Equal(97, stream.RequestedCounts[0]);
        Assert.All(stream.RequestedCounts.Skip(1), count => Assert.InRange(count, 1, 32));
        Assert.Equal(129, stream.BytesRead);
    }

    [Fact]
    public void CoordinatedWordRegistrationDoesNotApplyDosLimitToModernDocContent() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph(new string('M', 1_024));
        using var package = new MemoryStream();
        document.Save(package);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddWordAndLegacyHandlers(new LegacyWordImportOptions {
                Limits = new OfficeLegacyImportLimits { MaxInputBytes = 128 }
            })
            .Build();
        package.Position = 0;

        OfficeDocumentReadResult result = reader.ReadDocument(package, "modern.doc");
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("MMM", StringComparison.Ordinal));
    }

    [Fact]
    public void ReaderLimitCannotRaiseConfiguredLegacyWordLimit() {
        byte[] source = LegacyFixtureFactory.WordPerfect();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddLegacyWordHandler(new LegacyWordImportOptions {
                Limits = new OfficeLegacyImportLimits { MaxInputBytes = source.Length - 1 }
            })
            .Build();

        using var stream = new MemoryStream(source);
        Assert.Throws<IOException>(() => reader.ReadDocument(stream, "archive.wpd",
            new ReaderOptions { MaxInputBytes = source.Length + 100L }));
    }

    [Fact]
    public void LegacyWordHandlerUsesConfiguredLimitBeforeBufferingNonSeekableStreams() {
        byte[] source = LegacyFixtureFactory.WordPerfect().Concat(new byte[256 * 1024]).ToArray();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddLegacyWordHandler(new LegacyWordImportOptions {
                Limits = new OfficeLegacyImportLimits { MaxInputBytes = 128 }
            })
            .Build();
        using var stream = new NonSeekableStream(source);

        Assert.Throws<IOException>(() => reader.ReadDocument(stream, "archive.wpd",
            new ReaderOptions { MaxInputBytes = source.Length + 100L }));
        Assert.True(stream.BytesRead < source.Length);
    }

    [Fact]
    public void CompoundFamiliesRequireAValidCompoundDirectory() {
        Assert.Throws<InvalidDataException>(() => LegacyWordImporter.Import(
            LegacyFixtureFactory.TruncatedCompoundHeader(),
            new LegacyWordImportOptions { SourceName = "archive.lwp" }));
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

    private sealed class NonSeekableStream : Stream {
        private readonly MemoryStream _inner;

        internal NonSeekableStream(byte[] data) => _inner = new MemoryStream(data, writable: false);
        internal long BytesRead { get; private set; }
        internal List<int> RequestedCounts { get; } = new();
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) {
            RequestedCounts.Add(count);
            int read = _inner.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            RequestedCounts.Add(count);
            int read = await _inner.ReadAsync(buffer, offset, count, cancellationToken);
            BytesRead += read;
            return read;
        }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
