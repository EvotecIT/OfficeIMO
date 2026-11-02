using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfPdfNamedFontTests {
    private const string Source = "Portable Source";
    private const string Target = "Portable Substitute";

    [Theory]
    [InlineData("registered-source")]
    [InlineData("planned-source")]
    [InlineData("planned-default")]
    [InlineData("explicit-default")]
    public void CallerNamedRegistrationAndSubstitutionKeepTheirMeasuredGlyphs(string scenario) {
        RtfDocument document = CreateDocument(out int sourceId);
        if (scenario is "planned-default" or "explicit-default") document.Settings.SetDefaultFont(sourceId);
        RtfRun run = document.AddParagraph().AddText("AB");
        if (scenario != "planned-default") run.FontId = sourceId;
        run.FontSize = 10;
        string expectedFamily = scenario == "registered-source" ? Source : Target;
        PdfCore.PdfCreateOptions caller = CreateOptions(expectedFamily, substitute: scenario != "registered-source");
        if (scenario == "explicit-default") caller.DefaultFont = PdfCore.PdfStandardFont.Courier;
        byte[] sourceBefore = document.ToBytes();
        byte[] callerFont = caller.NamedFontFamilies[expectedFamily].Regular;

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(ConvertOptions(caller));
        using PdfPigDocument actual = PdfPigDocument.Open(result.ToBytes());
        Assert.Equal(1, actual.NumberOfPages);
        var word = Assert.Single(actual.GetPage(1).GetWords(), word => word.Text == "AB");
        Assert.All(word.Letters, letter => Assert.Contains(
            scenario == "explicit-default" ? "Courier" : expectedFamily.Replace(" ", string.Empty),
            letter.FontName, StringComparison.OrdinalIgnoreCase));
        Assert.Equal(scenario == "explicit-default" ? 6D : 5D,
            word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
        if (scenario is "planned-source" or "planned-default") {
            PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings,
                warning => warning.Details.TryGetValue("plannedSubstitution", out string? planned) && planned == bool.TrueString);
            Assert.Equal(Source, warning.Details["fontFamily"]);
            Assert.Equal(Target, warning.Details["resolvedFontFamily"]);
            Assert.Equal(PdfCore.PdfConversionWarningSeverity.Information, warning.Severity);
        }
        Assert.Equal(sourceBefore, document.ToBytes());
        Assert.Equal(callerFont, caller.NamedFontFamilies[expectedFamily].Regular);
        Assert.Empty(caller.EmbeddedFonts);
    }

    [Theory]
    [InlineData("usual-slot-known-control")]
    [InlineData("relocated-known")]
    [InlineData("relocated-custom")]
    [InlineData("literal-comma")]
    public void CallerEmbeddedFamilyRetainsItsRegisteredSlotForNonDefaultRuns(string scenario) {
        (string familyName, PdfCore.PdfStandardFont slot) = scenario switch {
            "usual-slot-known-control" => ("Times New Roman", PdfCore.PdfStandardFont.TimesRoman),
            "relocated-known" => ("Times New Roman", PdfCore.PdfStandardFont.Courier),
            "relocated-custom" => ("Portable Registered", PdfCore.PdfStandardFont.TimesRoman),
            "literal-comma" => ("Portable, Courier", PdfCore.PdfStandardFont.TimesRoman),
            _ => throw new ArgumentOutOfRangeException(nameof(scenario))
        };
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont("Helvetica"));
        RtfRun run = document.AddParagraph().AddText("AAAA");
        run.FontId = document.AddFont(familyName);
        run.FontSize = 12;
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300,
            Enumerable.Range(32, 95).ToArray());
        var caller = new PdfCore.PdfCreateOptions().RegisterFontFamily(slot,
            new PdfCore.PdfEmbeddedFontFamily(familyName, font));
        string resourceName = caller.EmbeddedFonts[slot].FontName!.Replace(" ", string.Empty).Replace(",", string.Empty);
        var embeddedBefore = caller.EmbeddedFonts.ToDictionary(entry => entry.Key,
            entry => new { entry.Value.FontName, Data = entry.Value.Data.ToArray() });
        byte[] sourceBefore = document.ToBytes();

        byte[] bytes = document.ConvertToPdfWithEvidence(ConvertOptions(caller)).ToBytes();
        using PdfPigDocument actual = PdfPigDocument.Open(bytes);
        Assert.Equal(1, actual.NumberOfPages);
        var letters = actual.GetPage(1).Letters;
        Assert.Equal(4, letters.Count);
        Assert.Equal("AAAA", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => {
            Assert.EndsWith(resourceName, letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(12D, letter.PointSize, 3);
        });
        var painted = PdfCore.PdfReadDocument.Open(bytes);
        var spans = Assert.Single(painted.Pages).GetTextSpans();
        var glyphs = spans.SelectMany(span => {
            Assert.True(span.TryGetGlyphs(out IReadOnlyList<PdfCore.PdfTextGlyph> values));
            return values;
        }).ToArray();
        Assert.Equal(4, glyphs.Length);
        Assert.All(glyphs, glyph => Assert.Equal(3.6D, glyph.PaintedAdvance, 3));
        Assert.Equal(sourceBefore, document.ToBytes());
        Assert.Equal(embeddedBefore.Keys.OrderBy(key => key), caller.EmbeddedFonts.Keys.OrderBy(key => key));
        foreach (var entry in embeddedBefore) {
            Assert.Equal(entry.Value.FontName, caller.EmbeddedFonts[entry.Key].FontName);
            Assert.Equal(entry.Value.Data, caller.EmbeddedFonts[entry.Key].Data);
        }
    }

    [Theory]
    [InlineData(false, false, 5D, "Regular")]
    [InlineData(true, false, 8D, "Bold")]
    [InlineData(false, true, 6D, "Italic")]
    [InlineData(true, true, 9D, "BoldItalic")]
    public void CallerNamedSubstitutionSelectsEachStyledFace(bool bold, bool italic, double advance, string face) {
        RtfDocument document = CreateDocument(out int sourceId);
        RtfRun run = document.AddParagraph().AddText("AB");
        run.FontId = sourceId;
        run.FontSize = 10;
        run.Bold = bold;
        run.Italic = italic;

        using PdfCore.PdfDocument actualModel = document.ToPdf(ConvertOptions(CreateOptions(Target, substitute: true)));
        using PdfPigDocument actual = PdfPigDocument.Open(actualModel.ToBytes());
        var letters = actual.GetPage(1).Letters;
        Assert.Equal("AB", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => Assert.EndsWith("PortableSubstitute-" + face, letter.FontName, StringComparison.OrdinalIgnoreCase));
        Assert.Equal(advance, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("table")]
    [InlineData("note")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("field")]
    [InlineData("inherited")]
    [InlineData("control-only")]
    [InlineData("empty")]
    public void InheritedNamedFontsReachEveryExistingRichProjection(string owner) {
        RtfDocument document = CreateDocument(out int sourceId);
        RtfStyle style = document.AddStyle(7, "Portable mono");
        style.FontId = sourceId;
        style.FontSize = 10;
        RtfParagraph paragraph;
        switch (owner) {
            case "table": paragraph = document.AddTable(1, 1).Rows[0].Cells[0].AddParagraph(); break;
            case "note":
                var note = new RtfNote(RtfNoteKind.Footnote);
                document.AddParagraph("BODY").AddNoteReference(note, "1");
                paragraph = note.AddParagraph();
                break;
            case "header":
                document.AddParagraph("BODY");
                paragraph = document.AddHeader().AddParagraph();
                break;
            case "footer":
                document.AddParagraph("BODY");
                paragraph = document.AddFooter().AddParagraph();
                break;
            case "field": paragraph = document.AddParagraph().AddField("MERGEFIELD sample").Result; break;
            default: paragraph = document.AddParagraph(); break;
        }
        paragraph.SetStyle(7);
        if (owner == "inherited") {
            document.AddStyle(8, "Child portable").BasedOnStyleId = 7;
            paragraph.SetStyle(8);
        }
        if (owner == "control-only") {
            paragraph.AddTabStop(720, RtfTabAlignment.Left, RtfTabLeader.Dots);
            paragraph.AddText("\t");
        } else if (owner != "empty") paragraph.AddText("AB");
        byte[] sourceBefore = document.ToBytes();
        PdfCore.PdfDocument converted = document.ToPdf(ConvertOptions(CreateOptions(Target, substitute: true)));
        if (owner == "empty") {
            PdfCore.RichParagraphBlock projected = Assert.IsType<PdfCore.RichParagraphBlock>(Assert.Single(converted.Blocks));
            Assert.Equal("\"" + Target + "\"", Assert.Single(projected.Runs, run => run.IsLineBoxStrut).FontFamily);
        }
        using PdfPigDocument actual = PdfPigDocument.Open(converted.ToBytes());
        Assert.Equal(1, actual.NumberOfPages);
        if (owner == "empty") Assert.Empty(actual.GetPage(1).Letters);
        else if (owner == "control-only") {
            var dots = actual.GetPage(1).Letters.Where(letter => letter.Value == ".").ToArray();
            Assert.True(dots.Length >= 3);
            Assert.All(dots, letter => Assert.Contains("PortableSubstitute", letter.FontName, StringComparison.OrdinalIgnoreCase));
        } else {
            var word = Assert.Single(actual.GetPage(1).GetWords(), word => word.Text == "AB");
            Assert.All(word.Letters, letter => Assert.Contains("PortableSubstitute", letter.FontName, StringComparison.OrdinalIgnoreCase));
            Assert.Equal(5D, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
        }
        Assert.Equal(sourceBefore, document.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DocumentDefaultFontAlsoReachesGeneratedMarkersListsAndNoteLabels(bool explicitDefault) {
        RtfDocument document = CreateDocument(out int sourceId);
        document.Settings.SetDefaultFont(sourceId);
        document.AddParagraph("AB");
        document.AddParagraph().AddPageNumber("CD");
        document.AddParagraph("EF").SetList(listId: 3, level: 0, kind: RtfListKind.Decimal).SetListText("7.\t");
        var note = new RtfNote(RtfNoteKind.Footnote);
        note.AddParagraph("GH");
        document.AddParagraph("IJ ").AddNoteReference(note, "5");
        PdfCore.PdfCreateOptions caller = CreateOptions(Target, substitute: true);
        if (explicitDefault) caller.DefaultFont = PdfCore.PdfStandardFont.Courier;
        byte[] sourceBefore = document.ToBytes();

        using PdfCore.PdfDocument actualModel = document.ToPdf(ConvertOptions(caller));
        using PdfPigDocument actual = PdfPigDocument.Open(actualModel.ToBytes());
        Assert.Equal(1, actual.NumberOfPages);
        Assert.All(actual.GetPage(1).Letters, letter => Assert.Contains(
            explicitDefault ? "Courier" : "PortableSubstitute", letter.FontName, StringComparison.OrdinalIgnoreCase));
        var words = actual.GetPage(1).GetWords().ToArray();
        foreach (string marker in new[] { "AB", "CD", "EF", "GH", "IJ", "7." }) {
            Assert.Single(words, word => word.Text == marker);
        }
        var label = Assert.Single(words, word => word.Text == "Footnote");
        Assert.All(label.Letters, letter => {
            Assert.Equal(9D, letter.PointSize, 3);
            Assert.Contains("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase);
        });
        Assert.Equal(explicitDefault ? 5.4D : 7.2D,
            label.Letters[1].StartBaseLine.X - label.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(sourceBefore, document.ToBytes());
    }

    [Theory]
    [InlineData(false, "plain")]
    [InlineData(true, "plain")]
    [InlineData(false, "literal-comma")]
    [InlineData(true, "literal-comma")]
    public void UnresolvedSecondaryFamilyUsesTheConfiguredDocumentDefault(bool explicitDefault, string scenario) {
        RtfDocument document = CreateDocument(out int sourceId);
        document.Settings.SetDefaultFont(sourceId);
        RtfRun run = document.AddParagraph().AddText("AB");
        string unavailableFamily = scenario == "literal-comma" ? "Unavailable, Courier" : "Unavailable Secondary";
        run.FontId = document.AddFont(unavailableFamily);
        run.FontSize = 10;
        PdfCore.PdfCreateOptions caller = CreateOptions(Target, substitute: true);
        if (explicitDefault) caller.DefaultFont = PdfCore.PdfStandardFont.Courier;
        byte[] sourceBefore = document.ToBytes();

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(ConvertOptions(caller));
        using PdfPigDocument actual = PdfPigDocument.Open(result.ToBytes());
        var word = Assert.Single(actual.GetPage(1).GetWords(), word => word.Text == "AB");
        Assert.All(word.Letters, letter => Assert.Contains(explicitDefault ? "Courier" : "PortableSubstitute",
            letter.FontName, StringComparison.OrdinalIgnoreCase));
        Assert.Equal(explicitDefault ? 6D : 5D,
            word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings,
            warning => warning.Details.TryGetValue("fontFamily", out string? source) && source == unavailableFamily);
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, warning.Severity);
        if (explicitDefault) Assert.Equal("Courier", warning.Details["fallbackSlot"]);
        else Assert.Equal(Target, warning.Details["resolvedFontFamily"]);
        Assert.Equal(sourceBefore, document.ToBytes());
    }

    [Theory]
    [InlineData("Portable, Source", "Portable, Source", false)]
    [InlineData(Source, "Portable, Substitute", true)]
    public void CallerLiteralNamedFamiliesRemainSingleFamilies(string sourceName, string targetName, bool substitute) {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont("Helvetica"));
        RtfRun run = document.AddParagraph().AddText("AB");
        run.FontId = document.AddFont(sourceName);
        run.FontSize = 10;
        PdfCore.PdfCreateOptions caller = CreateOptions(targetName, substitute: false);
        if (substitute) caller.RegisterFontFamilySubstitution(sourceName, targetName);

        using PdfCore.PdfDocument actualModel = document.ToPdf(ConvertOptions(caller));
        using PdfPigDocument actual = PdfPigDocument.Open(actualModel.ToBytes());
        var letters = actual.GetPage(1).Letters;
        Assert.Equal("AB", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => Assert.Contains("Portable", letter.FontName, StringComparison.OrdinalIgnoreCase));
        Assert.Equal(5D, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 3);
        Assert.True(caller.NamedFontFamilies.ContainsKey(targetName));
    }

    private static RtfDocument CreateDocument(out int sourceId) {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont("Helvetica"));
        sourceId = document.AddFont(Source);
        return document;
    }

    private static PdfCore.PdfCreateOptions CreateOptions(string familyName, bool substitute) {
        int[] scalars = Enumerable.Range(32, 95).ToArray();
        var family = new PdfCore.PdfEmbeddedFontFamily(familyName,
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, scalars),
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 800, scalars),
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, scalars),
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 900, scalars));
        var options = new PdfCore.PdfCreateOptions { DefaultFontSize = OfficeLength.Points(10) }.RegisterNamedFontFamily(family);
        if (substitute) options.RegisterFontFamilySubstitution(Source, familyName, PdfCore.PdfFontFamilySubstitutionImpact.Compatible);
        return options;
    }

    private static RtfToPdfOptions ConvertOptions(PdfCore.PdfCreateOptions options) => new RtfToPdfOptions {
        PdfOptions = options,
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
    };
}
