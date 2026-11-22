using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    // Word places a drop cap at the start of the following paragraph and indents that
    // paragraph's first lines by the cap's width, for as many lines as the cap spans.
    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void SaveAsPdf_DropCapIndentsTheFollowingLinesItSpans(int lines) {
        using WordDocument document = CreateDropCapDocument(lines, "drop", hSpace: 0);
        using PdfPigDocument pdf = RenderDropCapPdf(document);
        var cap = FindDropCapWord(pdf, "D");
        List<double> lineStarts = DropCapBodyLineStarts(pdf);

        Assert.Equal(DropCapMargin, cap.BoundingBox.Left, 0.5);
        Assert.All(lineStarts.Take(lines), start => Assert.Equal(cap.Letters[0].EndBaseLine.X, start, 0.5));
        Assert.Equal(DropCapMargin, lineStarts[lines], 0.15);
        Assert.False(document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false }).Warnings
            .Any(warning => warning.Code == "native-paragraph-frame-flattened"));
    }

    [Fact]
    public void SaveAsPdf_DropCapHorizontalSpaceWidensTheIndent() {
        using WordDocument plain = CreateDropCapDocument(3, "drop", hSpace: 0);
        using WordDocument spaced = CreateDropCapDocument(3, "drop", hSpace: 360);
        using PdfPigDocument plainPdf = RenderDropCapPdf(plain);
        using PdfPigDocument spacedPdf = RenderDropCapPdf(spaced);
        Assert.Equal(DropCapBodyLineStarts(plainPdf)[0] + 18D, DropCapBodyLineStarts(spacedPdf)[0], 0.15);
    }

    [Fact]
    public void SaveAsPdf_MarginDropCapHangsBesideTheText() {
        using WordDocument document = CreateDropCapDocument(3, "margin", hSpace: 0);
        using PdfPigDocument pdf = RenderDropCapPdf(document);
        Assert.Equal(DropCapMargin, FindDropCapWord(pdf, "D").Letters[0].EndBaseLine.X, 0.5);
        Assert.All(DropCapBodyLineStarts(pdf).Take(4), start => Assert.Equal(DropCapMargin, start, 0.15));
    }

    [Fact]
    public void SaveAsPdf_ShortParagraphBesideADropCapLetsTheNextOneWrap() {
        // Word wraps the paragraph after a short one around the rest of the cap.
        using WordDocument document = CreateDropCapDocument(3, "drop", hSpace: 0, bodyWords: 3);
        using PdfPigDocument pdf = RenderDropCapPdf(document);
        double indent = FindDropCapWord(pdf, "D").Letters[0].EndBaseLine.X;
        Assert.Equal(indent, FindDropCapWord(pdf, "AFTER").BoundingBox.Left, 0.5);
    }

    // w:position raises or lowers text by signed half-points; Word's drop caps use it to sit the cap on the
    // lines beside it. In a paragraph with exact line spacing the line box is unchanged.
    [Theory]
    [InlineData(24, 12D)]
    [InlineData(-8, -4D)]
    public void SaveAsPdf_PositionedTextMovesItsGlyphsFromTheBaseline(int halfPoints, double expectedRise) {
        double Baseline(int? position) {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("BEFORE ");
            paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.AddChild(new SpacingBetweenLines { Line = "340", LineRule = LineSpacingRuleValues.Exact }, true);
            var run = new Run(new Text("MOVED"));
            if (position.HasValue) run.PrependChild(new RunProperties(new Position { Val = position.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) }));
            paragraph._paragraph.Append(run);
            using PdfPigDocument pdf = RenderDropCapPdf(document);
            var words = pdf.GetPage(1).GetWords().ToList();
            return words.Single(word => word.Text.EndsWith("MOVED", StringComparison.Ordinal)).Letters.Last().StartBaseLine.Y -
                words.Single(word => word.Text.StartsWith("BEFORE", StringComparison.Ordinal)).Letters[0].StartBaseLine.Y;
        }
        Assert.Equal(0D, Baseline(null), 0.05);
        Assert.Equal(expectedRise, Baseline(halfPoints), 0.05);
    }

    private const double DropCapMargin = 72D;

    private static WordDocument CreateDropCapDocument(int lines, string kind, int hSpace, int bodyWords = 90) {
        WordDocument document = WordDocument.Create();
        document.Sections[0].PageSetup.Size.PageSize = WordPageSize.Letter;
        document.Sections[0].PageSetup.Margins.Left = document.Sections[0].PageSetup.Margins.Right = global::OfficeIMO.OfficeLength.Points(DropCapMargin);
        document.AddParagraph("Intro.");
        // The size, exact line height and lowering Word writes for each drop cap height.
        (int size, int line, int position) = lines switch { 2 => (77, 553, -4), 3 => (111, 827, -8), _ => (148, 1102, -10) };
        WordParagraph cap = document.AddParagraph();
        cap._paragraph.ParagraphProperties = new ParagraphProperties(
            new KeepNext(),
            new FrameProperties {
                DropCap = kind == "margin" ? DropCapLocationValues.Margin : DropCapLocationValues.Drop, Lines = lines,
                Wrap = TextWrappingValues.Around, VerticalPosition = VerticalAnchorValues.Text,
                HorizontalPosition = kind == "margin" ? HorizontalAnchorValues.Page : HorizontalAnchorValues.Text,
                HorizontalSpace = hSpace > 0 ? hSpace.ToString(System.Globalization.CultureInfo.InvariantCulture) : null
            },
            new SpacingBetweenLines { After = "0", Line = line.ToString(System.Globalization.CultureInfo.InvariantCulture), LineRule = LineSpacingRuleValues.Exact },
            new TextAlignment { Val = VerticalTextAlignmentValues.Baseline });
        cap._paragraph.Append(new Run(new RunProperties(
                new Position { Val = position.ToString(System.Globalization.CultureInfo.InvariantCulture) },
                new FontSize { Val = size.ToString(System.Globalization.CultureInfo.InvariantCulture) },
                new FontSizeComplexScript { Val = size.ToString(System.Globalization.CultureInfo.InvariantCulture) }),
            new Text("D")));
        document.AddParagraph(string.Join(" ", Enumerable.Range(1, bodyWords).Select(index => "w" + index.ToString("00", System.Globalization.CultureInfo.InvariantCulture))));
        document.AddParagraph("AFTER paragraph.");
        return document;
    }

    private static PdfPigDocument RenderDropCapPdf(WordDocument document) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        return PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
    }

    private static UglyToad.PdfPig.Content.Word FindDropCapWord(PdfPigDocument pdf, string token) =>
        pdf.GetPage(1).GetWords().First(word => word.Text == token);

    // Left edge of each line of the body paragraph, top to bottom.
    private static List<double> DropCapBodyLineStarts(PdfPigDocument pdf) =>
        pdf.GetPage(1).GetWords().Where(word => word.Text.Length == 3 && word.Text[0] == 'w' && char.IsDigit(word.Text[1]))
            .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 1))
            .OrderByDescending(line => line.Key)
            .Select(line => line.Min(word => word.BoundingBox.Left)).ToList();
}
