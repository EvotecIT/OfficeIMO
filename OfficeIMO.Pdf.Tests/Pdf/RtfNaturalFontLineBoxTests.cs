using OfficeIMO.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.TestAssets;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class RtfNaturalFontLineBoxTests {
    [Fact]
    public void ControlOnlyNoteBodyKeepsItsFontBoxesBeneathThePrependedLabel() {
        RtfDocument Create(bool reference) {
            RtfDocument source = CreateDocument(out int normal, out _);
            RtfStyle style = source.AddStyle(7, "Note body character box");
            style.FontId = normal;
            style.FontSize = 12D;
            style.LineSpacingTwips = 240;
            style.LineSpacingMultiple = true;
            var note = new RtfNote(RtfNoteKind.Footnote);
            RtfParagraph first = note.AddParagraph().SetStyle(style.Id);
            if (reference) note.AddParagraph().SetStyle(style.Id);
            else first.AddLineBreak();
            AddText(note.AddParagraph(), "Z", 12, normal);
            RtfParagraph body = source.AddParagraph();
            AddText(body, "X", 12, normal);
            body.AddNoteReference(note, "7");
            return source;
        }
        using PdfPigDocument expected = PdfPigDocument.Open((Create(reference: true).ToPdf(Options())).ToBytes());
        using PdfPigDocument actual = PdfPigDocument.Open((Create(reference: false).ToPdf(Options())).ToBytes());
        Letter[] referenceLetters = Markers(expected, "F", "Z");
        Letter[] letters = Markers(actual, "F", "Z");

        Assert.Equal(referenceLetters[0].StartBaseLine.Y, letters[0].StartBaseLine.Y, 3);
        Assert.Equal(referenceLetters[1].StartBaseLine.Y, letters[1].StartBaseLine.Y, 3);
        Assert.Equal(12D, letters[1].FontSize, 3);
    }

    [Fact]
    public void PlainLineBreakOnlyParagraphKeepsItsSourceDefaultFontBoxes() {
        const string prefix = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\plain A\par";
        RtfDocument actualSource = RtfDocument.Read(prefix + @"\line\par B\par}").Document;
        RtfDocument emptyBoxes = RtfDocument.Read(prefix + @"\par\par B\par}").Document;
        RtfToPdfOptions options = Options();
        Assert.IsType<PdfCreateOptions>(options.PdfOptions).DefaultFontSize = 64D;

        using PdfPigDocument expected = PdfPigDocument.Open((emptyBoxes.ToPdf(options)).ToBytes());
        using PdfPigDocument actual = PdfPigDocument.Open((actualSource.ToPdf(options)).ToBytes());
        Letter[] reference = Markers(expected, "A", "B");
        Letter[] letters = Markers(actual, "A", "B");

        // A line break creates two empty source lines before the paragraph ends.
        Assert.Equal(43.2D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(reference[1].StartBaseLine.Y, letters[1].StartBaseLine.Y, 3);
        Assert.All(letters, letter => Assert.Equal(12D, letter.FontSize, 3));
    }

    [Fact]
    public void PlainTabOnlyParagraphOccupiesTheSameFontBoxAsAnEmptyParagraph() {
        const string prefix = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\plain A\par";
        RtfDocument actualSource = RtfDocument.Read(prefix + @"\tab\par B\par}").Document;
        RtfDocument emptyBox = RtfDocument.Read(prefix + @"\par B\par}").Document;
        RtfToPdfOptions options = Options();
        Assert.IsType<PdfCreateOptions>(options.PdfOptions).DefaultFontSize = 64D;

        using PdfPigDocument expected = PdfPigDocument.Open((emptyBox.ToPdf(options)).ToBytes());
        using PdfPigDocument actual = PdfPigDocument.Open((actualSource.ToPdf(options)).ToBytes());
        Letter[] reference = Markers(expected, "A", "B");
        Letter[] letters = Markers(actual, "A", "B");

        Assert.Equal(28.8D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(reference[1].StartBaseLine.Y, letters[1].StartBaseLine.Y, 3);
        Assert.Equal(2, actual.GetPage(1).Letters.Count);
    }

    [Theory]
    [InlineData("line", 2, 72D)]
    [InlineData("tab", 1, 43.2D)]
    public void ControlOnlyParagraphKeepsItsAvailableInheritedCharacterSize(
        string control, int emptyBoxCount, double expectedGap) {
        RtfDocument Create(bool reference) {
            RtfDocument source = CreateDocument(out int normal, out _);
            RtfStyle style = source.AddStyle(7, "Large control-only paragraph");
            style.FontId = normal;
            style.FontSize = 24D;
            style.LineSpacingTwips = 240;
            style.LineSpacingMultiple = true;
            AddText(source.AddParagraph(), "A", 12, normal);
            for (int index = 0; index < (reference ? emptyBoxCount : 1); index++) {
                RtfParagraph paragraph = source.AddParagraph().SetStyle(style.Id);
                if (!reference) {
                    if (control == "line") paragraph.AddLineBreak();
                    else paragraph.AddText("\t");
                }
            }
            AddText(source.AddParagraph(), "B", 12, normal);
            return source;
        }
        RtfToPdfOptions options = Options();
        Assert.IsType<PdfCreateOptions>(options.PdfOptions).DefaultFontSize = 64D;
        using PdfPigDocument expected = PdfPigDocument.Open((Create(reference: true).ToPdf(options)).ToBytes());
        using PdfPigDocument actual = PdfPigDocument.Open((Create(reference: false).ToPdf(options)).ToBytes());
        Letter[] referenceLetters = Markers(expected, "A", "B");
        Letter[] letters = Markers(actual, "A", "B");

        Assert.Equal(expectedGap, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(referenceLetters[1].StartBaseLine.Y, letters[1].StartBaseLine.Y, 3);
        Assert.Equal(2, actual.GetPage(1).Letters.Count);
    }

    [Fact]
    public void TabOnlyParagraphUsesItsAvailableDirectRunSize() {
        RtfDocument source = CreateDocument(out int normal, out _);
        AddText(source.AddParagraph(), "A", 12, normal);
        RtfParagraph paragraph = source.AddParagraph().SetLineSpacing(240, multiple: true);
        AddText(paragraph, "\t", 24, normal);
        AddText(source.AddParagraph(), "B", 12, normal);
        RtfToPdfOptions options = Options();
        Assert.IsType<PdfCreateOptions>(options.PdfOptions).DefaultFontSize = 64D;

        using PdfPigDocument pdf = PdfPigDocument.Open((source.ToPdf(options)).ToBytes());
        Letter[] letters = Markers(pdf, "A", "B");

        Assert.Equal(43.2D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(2, pdf.GetPage(1).Letters.Count);
    }

    [Fact]
    public void SelectedCffProgramSuppliesNaturalRtfLineSpacing() {
        string path = Assert.IsType<string>(PdfComplianceTestFonts.FindBundledOpenTypeCffFont());
        byte[] font = File.ReadAllBytes(path);
        RtfDocument source = CreateDocument(out int normal, out _);
        source.PageSetup.SetPaperSize(12240, 15840);
        source.PageSetup.SetMargins(leftTwips: 1440, rightTwips: 1440, topTwips: 1440, bottomTwips: 1440);
        RtfParagraph paragraph = source.AddParagraph().SetLineSpacing(360, multiple: true);
        AddLine(paragraph, "A", 12, normal);
        AddText(paragraph, "B", 12, normal);
        var options = new PdfCreateOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D),
            CompressContentStreams = false
        }.EmbedStandardFont(PdfStandardFont.Helvetica, font, "RTF Source Serif CFF");

        byte[] bytes = (source.ToPdf(new RtfToPdfOptions {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            TextFallbacks = PdfTextFallbackFeatures.None,
            PdfOptions = options
        })).ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Letter[] letters = Markers(pdf, "A", "B");

        // The bundled Source Serif CFF stores 1000 units/em, ascent/descent
        // 1036/-335 and no gap. Expectations come from that fixture's tables,
        // independently of the production metrics lookup.
        Assert.Equal(707.568D, letters[0].StartBaseLine.Y, 3);
        Assert.Equal(24.678D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.All(letters, letter => Assert.Equal(12D, letter.FontSize, 3));
        string raw = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/FontFile3", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType0C", raw, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(12D, 12D)]
    [InlineData(12D, 18D)]
    [InlineData(12D, 30D)]
    [InlineData(24D, 18D)]
    public void ExactLineBaselineUsesTheAuthoredBoxIndependentlyOfFontMetrics(double fontSize, double height) {
        foreach (bool useTallFont in new[] { false, true }) {
            RtfDocument source = CreateDocument(out int normal, out int tall);
            source.PageSetup.SetPaperSize(12240, 15840);
            source.PageSetup.SetMargins(leftTwips: 1440, rightTwips: 1440, topTwips: 1440, bottomTwips: 1440);
            RtfParagraph paragraph = source.AddParagraph().SetLineSpacing(-(int)(height * 20D), multiple: false);
            int font = useTallFont ? tall : normal;
            AddLine(paragraph, "A", fontSize, font);
            AddText(paragraph, "B", fontSize, font);

            using PdfPigDocument pdf = PdfPigDocument.Open((source.ToPdf(Options())).ToBytes());
            Letter[] letters = Markers(pdf, "A", "B");

            Assert.Equal(720D - height * 4D / 5D, letters[0].StartBaseLine.Y, 3);
            Assert.Equal(height, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
            Assert.All(letters, letter => Assert.Equal(fontSize, letter.FontSize, 3));
            // Exact spacing preserves its advance even when larger glyphs
            // extend beyond or overlap the authored line boxes.
            Assert.Equal(2, pdf.GetPage(1).Letters.Count);
        }
    }

    [Theory]
    [InlineData(false, 40.8D, 25.2D)]
    [InlineData(true, 48D, 45.6D)]
    public void SeparateLinesUseTheirOwnFontProgramsInsteadOfTheTallestParagraphRatio(
        bool multiple, double firstGap, double secondGap) {
        RtfDocument source = CreateDocument(out int normal, out int tall);
        RtfParagraph paragraph = source.AddParagraph().SetLineSpacing(360, multiple);
        AddLine(paragraph, "A", 12, normal);
        AddLine(paragraph, "B", 24, tall);
        AddText(paragraph, "C", 12, normal);

        using PdfPigDocument pdf = PdfPigDocument.Open((source.ToPdf(Options())).ToBytes());
        Letter[] letters = Markers(pdf, "A", "B", "C");

        Assert.Equal(12D, letters[0].FontSize, 3);
        Assert.Equal(24D, letters[1].FontSize, 3);
        Assert.Equal(12D, letters[2].FontSize, 3);
        // Normal: (800 + 200 + 200) / 1000. Tall: (1600 + 400) / 1000.
        // The expected gaps include each following line's own baseline offset.
        Assert.Equal(firstGap, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(secondGap, letters[1].StartBaseLine.Y - letters[2].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("note")]
    public void NaturalMultipleHasTheSameMeaningInBodyCellAndNoteParagraphs(string context) {
        RtfDocument source = CreateDocument(out int normal, out _);
        RtfParagraph paragraph;
        if (context == "cell") {
            paragraph = source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph();
        } else if (context == "note") {
            var note = new RtfNote(RtfNoteKind.Footnote);
            paragraph = note.AddParagraph();
            source.AddParagraph("body").AddNoteReference(note, "7");
        } else {
            paragraph = source.AddParagraph();
        }
        paragraph.SetLineSpacing(360, multiple: true);
        AddLine(paragraph, "X", 12, normal);
        AddLine(paragraph, "Y", 24, normal);
        AddText(paragraph, "Z", 12, normal);

        using PdfPigDocument pdf = PdfPigDocument.Open((source.ToPdf(Options())).ToBytes());
        Letter[] letters = Markers(pdf, "X", "Y", "Z");

        Assert.Equal(33.6D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(31.2D, letters[1].StartBaseLine.Y - letters[2].StartBaseLine.Y, 3);
        if (context == "note") Assert.Contains("Footnote 7:", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void AbsentAfterSpacingDoesNotAddThePdfBuildersConvenienceGap() {
        RtfDocument source = RtfDocument.Read(
            @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\f0\fs24 A\par B\par C\par}").Document;

        using PdfPigDocument pdf = PdfPigDocument.Open((source.ToPdf(Options())).ToBytes());
        Letter[] letters = Markers(pdf, "A", "B", "C");

        Assert.Equal(14.4D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(14.4D, letters[1].StartBaseLine.Y - letters[2].StartBaseLine.Y, 3);
        Assert.All(source.Paragraphs, paragraph => Assert.Null(paragraph.SpaceAfterTwips));
    }

    [Fact]
    public void PlainResetResolvesTheSourceCharacterSizeWithoutRequiringAStyleTable() {
        RtfDocument source = RtfDocument.Read(
            @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\plain A\par B\par}").Document;
        RtfToPdfOptions options = Options();
        Assert.IsType<PdfCreateOptions>(options.PdfOptions).DefaultFontSize = 64D;

        using PdfPigDocument pdf = PdfPigDocument.Open((source.ToPdf(options)).ToBytes());
        Letter[] letters = Markers(pdf, "A", "B");

        Assert.All(letters, letter => Assert.Equal(12D, letter.FontSize, 3));
        Assert.Equal(14.4D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Empty(source.Styles);
    }

    [Fact]
    public void EmptyParagraphRetainsItsExplicitInheritedCharacterBox() {
        RtfDocument source = CreateDocument(out int normal, out _);
        RtfStyle mark = source.AddStyle(7, "Large empty paragraph");
        mark.FontId = normal;
        mark.FontSize = 24D;
        AddText(source.AddParagraph(), "A", 12, normal);
        source.AddParagraph().SetStyle(mark.Id);
        AddText(source.AddParagraph(), "B", 12, normal);

        using PdfPigDocument pdf = PdfPigDocument.Open((source.ToPdf(Options())).ToBytes());
        Letter[] letters = Markers(pdf, "A", "B");

        Assert.Equal(43.2D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(2, pdf.GetPage(1).Letters.Count);
    }

    [Fact]
    public void RaisedBodyRunsDoNotAcquireAnUnscaledDefaultParagraphStrut() {
        RtfDocument source = CreateDocument(out int normal, out _);
        RtfParagraph paragraph = source.AddParagraph().SetLineSpacing(240, multiple: true);
        AddText(paragraph, "A", 24, normal).VerticalPosition = RtfVerticalPosition.Superscript;
        paragraph.AddLineBreak();
        AddText(paragraph, "B", 24, normal).VerticalPosition = RtfVerticalPosition.Superscript;
        RtfToPdfOptions smallDefault = Options();
        RtfToPdfOptions largeDefault = Options();
        Assert.IsType<PdfCreateOptions>(largeDefault.PdfOptions).DefaultFontSize = 64D;

        using PdfPigDocument expected = PdfPigDocument.Open((source.ToPdf(smallDefault)).ToBytes());
        using PdfPigDocument actual = PdfPigDocument.Open((source.ToPdf(largeDefault)).ToBytes());
        Letter[] first = Markers(expected, "A", "B");
        Letter[] second = Markers(actual, "A", "B");

        for (int index = 0; index < first.Length; index++) {
            Assert.Equal(15.6D, second[index].FontSize, 3);
            Assert.Equal(first[index].StartBaseLine.Y, second[index].StartBaseLine.Y, 3);
        }
        Assert.True(second[0].StartBaseLine.Y > second[1].StartBaseLine.Y);
    }

    [Fact]
    public void ImplicitCellRunKeepsTheCallersTableFontSize() {
        RtfDocument source = CreateDocument(out int normal, out _);
        RtfParagraph paragraph = source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph()
            .SetLineSpacing(240, multiple: true);
        AddLine(paragraph, "A", 4, normal);
        paragraph.AddText("B").FontId = normal;
        RtfToPdfOptions options = Options();
        PdfCreateOptions configured = Assert.IsType<PdfCreateOptions>(options.PdfOptions);
        configured.DefaultFontSize = global::OfficeIMO.OfficeLength.Points(40D);
        configured.DefaultTableStyle = new PdfTableStyle { FontSize = 9D };

        using PdfPigDocument pdf = PdfPigDocument.Open((source.ToPdf(options)).ToBytes());
        Letter[] letters = Markers(pdf, "A", "B");

        Assert.Equal(4D, letters[0].FontSize, 3);
        Assert.Equal(9D, letters[1].FontSize, 3);
        Assert.Equal(40D, configured.DefaultFontSize.ToPoints());
        Assert.Equal(9D, Assert.IsType<PdfTableStyle>(configured.DefaultTableStyle).FontSize);
    }

    [Fact]
    public void MissingEmbeddedMetricsPreserveTheGenericPdfSpacingPolicy() {
        RtfDocument source = CreateDocument(out int normal, out _);
        RtfParagraph paragraph = source.AddParagraph().SetLineSpacing(360, multiple: true);
        AddLine(paragraph, "A", 12, normal);
        AddLine(paragraph, "B", 24, normal);
        AddText(paragraph, "C", 12, normal);
        var options = new RtfToPdfOptions {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            TextFallbacks = PdfTextFallbackFeatures.None,
            PdfOptions = new PdfCreateOptions { DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D) }
        };

        using PdfPigDocument pdf = PdfPigDocument.Open((source.ToPdf(options)).ToBytes());
        Letter[] letters = Markers(pdf, "A", "B", "C");

        Assert.Equal(54D, letters[0].StartBaseLine.Y - letters[2].StartBaseLine.Y, 3);
        Assert.Equal(12D, letters[0].FontSize, 3);
        Assert.Equal(24D, letters[1].FontSize, 3);
        Assert.Empty(Assert.IsType<PdfCreateOptions>(options.PdfOptions).EmbeddedFonts);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FallbackMetricsUseThePaintedProgramAndPreserveSectionResourceOwnership(bool section) {
        RtfDocument source = CreateDocument(out int normal, out _);
        RtfParagraph paragraph;
        if (section) {
            AddText(source.AddSection().AddParagraph(), "P", 12, normal);
            paragraph = source.AddSection(RtfSectionBreakKind.NextPage).AddParagraph();
        } else {
            paragraph = source.AddParagraph();
        }
        paragraph.SetLineSpacing(240, multiple: true);
        AddLine(paragraph, "A", 12, normal);
        AddLine(paragraph, "\u03a9", 12, normal);
        AddText(paragraph, "B", 12, normal);
        RtfToPdfOptions options = Options(includeTall: false);
        PdfCreateOptions configured = Assert.IsType<PdfCreateOptions>(options.PdfOptions);
        byte[] fallback = ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(1300, -300, 200, 300,
            ' ', 0x03a9);
        // Its requested compatibility slot is occupied by the caller's Arial.
        // The canonical fallback planner must select a different resource.
        configured.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("RTF Greek fallback", fallback)
        }, new[] { PdfStandardFont.Helvetica }));

        PdfDocumentConversionResult result = source.ConvertToPdfWithEvidence(options);
        Assert.Equal(1, EmbeddedFamilyCount(configured));
        if (section) {
            PageBlock page = Assert.Single(result.Value.Blocks.OfType<PageBlock>());
            Assert.Equal(2, EmbeddedFamilyCount(page.Options));
            Assert.Equal(1, EmbeddedFamilyCount(result.Value.Options));
        }
        byte[] bytes = result.Value.ToBytes();
        Assert.Equal(bytes, result.Value.ToBytes());
        PdfReadDocument read = PdfReadDocument.Open(bytes);
        var spans = read.Pages[section ? 1 : 0].GetTextSpans()
            .Where(span => span.Text.Length > 0).ToArray();
        PdfTextSpan omega = Assert.Single(spans, span => span.Text == "\u03a9");
        PdfTextSpan first = Assert.Single(spans, span => span.Text == "A");
        PdfTextSpan last = Assert.Single(spans, span => span.Text == "B");

        Assert.NotEqual(first.FontResource, omega.FontResource);
        Assert.Equal(first.FontResource, last.FontResource);
        Assert.Contains("RTFGreekfallback", Assert.IsType<string>(omega.BaseFont).Replace(" ", string.Empty), StringComparison.Ordinal);
        Assert.Equal(36D, first.Y - last.Y, 3);
        Assert.Equal(1, EmbeddedFamilyCount(configured));
    }

    private static RtfDocument CreateDocument(out int normal, out int tall) {
        RtfDocument document = RtfDocument.Create();
        normal = document.AddFont("Arial");
        tall = document.AddFont("Times New Roman");
        document.Settings.SetDefaultFont(normal);
        return document;
    }

    private static RtfRun AddText(RtfParagraph paragraph, string text, double size, int font) {
        RtfRun run = paragraph.AddText(text).SetFontSize(size);
        run.FontId = font;
        return run;
    }

    private static void AddLine(RtfParagraph paragraph, string text, double size, int font) {
        AddText(paragraph, text, size, font);
        paragraph.AddLineBreak();
    }

    private static Letter[] Markers(PdfPigDocument pdf, params string[] markers) {
        Assert.Equal(1, pdf.NumberOfPages);
        return markers.Select(marker => Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == marker)).ToArray();
    }

    private static int EmbeddedFamilyCount(PdfCreateOptions options) => options.EmbeddedFonts.Keys
        .Select(PdfStandardFontMapper.GetFontFamily).Distinct().Count();

    private static RtfToPdfOptions Options(bool includeTall = true) {
        int[] scalars = Enumerable.Range(32, 95).ToArray();
        var options = new PdfCreateOptions { DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D) };
        options.RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Arial",
            ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -200, 200, 200, scalars)));
        if (includeTall) options.RegisterFontFamily(PdfStandardFont.TimesRoman, new PdfEmbeddedFontFamily("Times New Roman",
            ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(1600, -400, 0, 400, scalars)));
        return new RtfToPdfOptions {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            TextFallbacks = PdfTextFallbackFeatures.None,
            PdfOptions = options
        };
    }
}
