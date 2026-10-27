using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using UglyToad.PdfPig.Content;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfPdfKerningRegressionTests {
    [Theory]
    [InlineData("body")]
    [InlineData("field")]
    [InlineData("cell")]
    [InlineData("note")]
    public void MissingKerningIsDisabledAndExplicitThresholdReachesRichTextStories(string route) {
        AssertPair(Assert.Single(ReadPairs(CreatePair(route, null), route + "-default")), kerned: false);
        AssertPair(Assert.Single(ReadPairs(CreatePair(route, 24), route + "-threshold")), kerned: true);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void KerningThresholdUsesInheritedOrPdfDefaultFontSize(bool styleFontSize) {
        foreach (int threshold in new[] { 26, 28 }) {
            RtfDocument source = RtfDocument.Create();
            RtfParagraph paragraph = source.AddParagraph();
            if (styleFontSize) {
                source.AddStyle(1, "Inherited size").FontSize = 13;
                paragraph.SetStyle(1);
            }
            RtfRun run = paragraph.AddText("TA").SetKerningHalfPoints(threshold);
            Assert.Null(run.FontSize);
            var options = PortableOptions(styleFontSize ? 12 : 13);

            var pair = Assert.Single(ReadPairs(source,
                "size-" + styleFontSize + "-" + threshold, options));

            Assert.Equal(13D, pair.Left.FontSize, 3);
            Assert.Equal(13D, pair.Right.FontSize, 3);
            AssertPair(pair, kerned: threshold == 26);
        }
    }

    [Fact]
    public void ParsedKerningStateInheritsIntoGroupsAndExplicitZeroDisablesIt() {
        const string rtf = @"{\rtf1\ansi\fs24\kerning24\pard {TA}\par\kerning0 TA\par}";
        RtfReadResult parsed = RtfDocument.Read(rtf);
        RtfParagraph[] paragraphs = parsed.Document.Paragraphs.ToArray();
        Assert.Equal(2, paragraphs.Length);
        Assert.Equal(24, Assert.Single(paragraphs[0].Runs).KerningHalfPoints);
        Assert.Null(Assert.Single(paragraphs[1].Runs).KerningHalfPoints);

        var pairs = ReadPairs(parsed.Document, "parsed-inherit-zero");

        Assert.Equal(2, pairs.Count);
        AssertPair(pairs[0], kerned: true);
        AssertPair(pairs[1], kerned: false);
        Assert.Equal(rtf, parsed.ToRtfLossless());
    }

    [Fact]
    public void GeneratedFallbackTextAndAppendedNoteLabelsUseUnkernedDefaults() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph().AddGeneratedText(RtfGeneratedTextKind.PageNumber, "TA");
        var note = new RtfNote(RtfNoteKind.Annotation) { Author = "TA" };
        note.AddParagraph("Z");
        source.AddParagraph("X").Runs[0].SetNote(note);

        var pairs = ReadPairs(source, "generated-and-label");

        Assert.Equal(2, pairs.Count);
        Assert.All(pairs, pair => AssertPair(pair, kerned: false));
    }

    private static RtfDocument CreatePair(string route, int? threshold) {
        RtfDocument source = RtfDocument.Create();
        RtfParagraph paragraph;
        switch (route) {
            case "body":
                paragraph = source.AddParagraph();
                break;
            case "field":
                paragraph = source.AddParagraph().AddField("QUOTE").Result;
                break;
            case "cell":
                paragraph = source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph();
                break;
            case "note":
                var note = new RtfNote(RtfNoteKind.Footnote);
                paragraph = note.AddParagraph();
                source.AddParagraph("X").Runs[0].SetNote(note);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(route));
        }
        RtfRun run = paragraph.AddText("TA").SetFontSize(12);
        if (threshold.HasValue) run.SetKerningHalfPoints(threshold.Value);
        return source;
    }

    private static IReadOnlyList<(Letter Left, Letter Right)> ReadPairs(RtfDocument source, string artifactName,
        RtfToPdfOptions? options = null) {
        string before = source.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfToPdfOptions caller = options ?? PortableOptions(12);
        PdfCore.PdfCreateOptions callerPdfOptions = caller.PdfOptions!;
        Assert.True(callerPdfOptions.TextKerning);
        PdfCore.PdfDocumentConversionResult result = source.ConvertToPdfWithEvidence(caller);
        Assert.True(callerPdfOptions.TextKerning);
        byte[] bytes = result.ToBytes();
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is { Length: > 0 } output) {
            Directory.CreateDirectory(output);
            File.WriteAllBytes(Path.Combine(output, "rtf-kerning-" + artifactName + ".pdf"), bytes);
        }
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Letter[] letters = pdf.GetPage(1).Letters.ToArray();
        var pairs = new List<(Letter Left, Letter Right)>();
        for (int index = 0; index + 1 < letters.Length; index++) {
            if (letters[index].Value == "T" && letters[index + 1].Value == "A") {
                pairs.Add((letters[index], letters[index + 1]));
            }
        }
        Assert.Equal(before, source.ToRtf(new RtfWriteOptions { IncludeGenerator = false }));
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "RunFormattingFlattened" &&
            warning.Details.TryGetValue("Features", out string? features) &&
            features?.Split(',').Contains("character-metrics") == true);
        return pairs;
    }

    private static void AssertPair((Letter Left, Letter Right) pair, bool kerned) {
        Assert.Equal(pair.Left.StartBaseLine.Y, pair.Right.StartBaseLine.Y, 3);
        if (kerned) {
            Assert.True(pair.Right.StartBaseLine.X < pair.Left.EndBaseLine.X);
        } else {
            Assert.Equal(pair.Left.EndBaseLine.X, pair.Right.StartBaseLine.X, 3);
        }
    }

    private static RtfToPdfOptions PortableOptions(double fontSize) => new RtfToPdfOptions {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        PdfOptions = new PdfCore.PdfCreateOptions {
            DefaultFont = PdfCore.PdfStandardFont.Helvetica,
            DefaultFontSize = fontSize,
            TextKerning = true
        }
    };
}
