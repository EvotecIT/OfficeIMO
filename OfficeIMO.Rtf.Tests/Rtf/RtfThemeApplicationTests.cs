using OfficeIMO.Drawing;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfThemeApplicationTests {
    [Fact]
    public void ExplicitThemeRestylesParsedContentWithoutReplacingItsSemanticsOrGeometry() {
        byte[] picture = OfficePngWriter.EncodeRgba(1, 1, new byte[] { 255, 0, 0, 255 });
        using RtfDocument document = RtfDocument.ParseResult(Source.Replace("PICTURE", BitConverter.ToString(picture).Replace("-", string.Empty))).Document;
        RtfParagraph body = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText().Contains("PLAIN", StringComparison.Ordinal));
        RtfRun plain = Run(document, "PLAIN");
        RtfRun code = Run(document, "CODE");
        RtfRun symbol = Run(document, "SYMBOL");
        RtfField link = Assert.Single(body.Inlines.OfType<RtfField>());
        RtfNote note = Assert.Single(document.Notes);
        RtfImage image = Assert.Single(document.Blocks.OfType<RtfImage>());
        RtfParagraph listParagraph = Assert.Single(document.Paragraphs, paragraph => paragraph.ListText != null);
        RtfParagraph marker = listParagraph.ListText!;
        RtfRun[] markerRuns = marker.Runs.ToArray();
        Assert.NotEmpty(markerRuns);
        Assert.All(markerRuns, run => Assert.Equal(2, run.FontId));
        RtfTable[] tables = document.Blocks.OfType<RtfTable>().ToArray();
        Assert.Equal(2, tables.Length);
        IRtfBlock[] blocks = document.Blocks.ToArray();
        string[] text = document.Paragraphs.Select(paragraph => paragraph.ToPlainText()).ToArray();
        Assert.Equal(9D, document.ResolveRunFormatting(body, plain).FontSize);
        string original = document.ToRtf();
        using (RtfDocument faithful = RtfDocument.ParseResult(original).Document)
            Assert.Equal(9D, faithful.ResolveRunFormatting(faithful.Paragraphs.Single(paragraph => paragraph.ToPlainText().Contains("PLAIN", StringComparison.Ordinal)), Run(faithful, "PLAIN")).FontSize);

        OfficeTheme theme = CustomTheme();
        Assert.Same(document, document.ApplyTheme(theme));
        Assert.Equal(blocks, document.Blocks);
        Assert.Equal(text, document.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
        Assert.Same(link, Assert.Single(body.Inlines.OfType<RtfField>()));
        Assert.Equal("HYPERLINK \"https://example.com/\"", link.Instruction);
        Assert.Same(note, Assert.Single(document.Notes));
        Assert.Contains("NOTEBODY", note.ToPlainText(), StringComparison.Ordinal);
        Assert.Same(marker, listParagraph.ListText);
        Assert.Equal(markerRuns, marker.Runs);
        Assert.All(marker.Runs, run => Assert.Equal(2, run.FontId));
        Assert.Equal(picture, image.Data);
        Assert.Equal(180, image.DesiredWidthTwips);
        Assert.Equal(240, image.DesiredHeightTwips);
        Assert.Equal((12240, 15840, 900, 1000, 1100, 1200), (document.PageSetup.PaperWidthTwips,
            document.PageSetup.PaperHeightTwips, document.PageSetup.MarginLeftTwips, document.PageSetup.MarginRightTwips,
            document.PageSetup.MarginTopTwips, document.PageSetup.MarginBottomTwips));
        Assert.Equal(RtfTextDirection.LeftToRight, body.Direction);
        Assert.Equal(1033, plain.LanguageId);
        Assert.True(Run(document, "BOLD").Bold);
        Assert.True(document.ResolveRunFormatting(body, Run(document, "ITALIC")).Italic);
        Assert.Equal("Courier New", Font(document, code));
        Assert.Equal("Symbol", Font(document, symbol));
        Assert.Equal("Arial", Font(document, plain));
        Assert.Equal(23D, Run(document, "HEADING").FontSize);
        Assert.Equal(11D, Run(document, "DEEPEST").FontSize);
        Assert.Equal(8, document.ResolveParagraphFormatting(document.Paragraphs.Single(paragraph => paragraph.ToPlainText() == "DEEPEST")).OutlineLevel);
        Assert.Equal(13D, plain.FontSize);
        Assert.Equal(140, body.SpaceAfterTwips);
        Assert.True(tables[0].Rows[0].RepeatHeader);
        Assert.False(tables[1].Rows[0].RepeatHeader);
        Assert.Equal(-600, tables[0].Rows[1].HeightTwips);
        Assert.Equal(RtfTableAlignment.Right, tables[0].Rows[1].Alignment);
        Assert.Equal(new int?[] { 2500, 5000 }, tables[0].Rows[1].Cells.Select(cell => cell.RightBoundaryTwips));
        Assert.Equal(RtfTableCellMerge.First, tables[0].Rows[1].Cells[0].HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, tables[0].Rows[1].Cells[1].HorizontalMerge);
        Assert.Equal(RtfTextAlignment.Right, tables[0].Rows[1].Cells[0].Paragraphs[0].Alignment);
        Assert.Equal(80, tables[0].Rows[0].Cells[0].PaddingLeftTwips);
        Assert.Equal((byte)255, document.GetColor(tables[0].Rows[0].Cells[0].BackgroundColorIndex!.Value)!.Blue);
        Assert.Null(tables[1].Rows[0].Cells[0].BackgroundColorIndex);
        Assert.False(tables[1].Rows[0].Cells[0].Paragraphs[0].Runs[0].Bold);

        var save = new RtfSaveOptions { IncludeGenerator = false };
        string themed = document.ToRtf(save);
        int fonts = document.Fonts.Count, colors = document.Colors.Count;
        document.ApplyTheme(theme);
        Assert.Equal(themed, document.ToRtf(save));
        Assert.Equal(fonts, document.Fonts.Count);
        Assert.Equal(colors, document.Colors.Count);
        using RtfDocument reopened = RtfDocument.ParseResult(themed).Document;
        Assert.Equal(13D, Run(reopened, "PLAIN").FontSize);
        Assert.Equal("Courier New", Font(reopened, Run(reopened, "CODE")));
        Assert.Equal(picture, Assert.Single(reopened.Blocks.OfType<RtfImage>()).Data);
        Assert.Equal(link.Instruction, Assert.Single(reopened.Paragraphs.SelectMany(paragraph => paragraph.Inlines).OfType<RtfField>()).Instruction);
        Assert.Contains("NOTEBODY", Assert.Single(reopened.Notes).ToPlainText(), StringComparison.Ordinal);
        using PdfCore.PdfDocument pdfModel = reopened.ToPdf(PortableOptions());
        using PdfPigDocument pdf = PdfPigDocument.Open(pdfModel.ToBytes());
        AssertPaint(pdf, "HEADING", 23D, "Times");
        AssertPaint(pdf, "PLAIN", 13D, "Helvetica");
        AssertPaint(pdf, "CODE", 13D, "Courier");
        Assert.Contains(pdf.GetPages().SelectMany(page => page.Paths), path => path.IsFilled && path.FillColor is { } fill && fill.ToRGBValues() == (0D, 0D, 1D));

        OfficeTheme mono = theme.WithFonts(theme.Fonts.With(body: theme.Fonts.Body.With(family: "Courier New")))
            .WithTextStyle(theme.TextStyle.With(headingBold: false));
        document.ApplyTheme(mono);
        Assert.Equal("Courier New", Font(document, plain));
        Assert.False(Run(document, "HEADING").Bold);
        Assert.True(Run(document, "BOLD").Bold);
        document.ApplyTheme(theme);
        Assert.Equal(themed, document.ToRtf(save));
        document.ApplyTheme(mono);
        RtfRun appended = body.AddText("APPENDED");
        document.ApplyTheme(theme);
        Assert.Equal("Arial", Font(document, appended));
        Assert.Equal("Courier New", Font(document, code));
        Assert.Equal("Symbol", Font(document, symbol));
    }

    [Fact]
    public void GeneratedAndEmptyParagraphsUseAuthoredStyleDefaultsBeforeAndAfterExplicitTheming() {
        using RtfDocument document = RtfDocument.ParseResult(@"{\rtf1\ansi\deff0{\fonttbl{\f0\fmodern Courier New;}}" +
            @"{\stylesheet{\s0\f0\fs36 Normal;}{\s1\sbasedon0 Derived;}}\pard\s1\chpgn\par\pard\s1\par\pard\s0 END\par}").Document;
        Assert.Null(document.Styles.Single(style => style.Id == 1 && style.Kind == RtfStyleKind.Paragraph).FontId);
        RtfGeneratedText generated = Assert.Single(document.Paragraphs[0].Inlines.OfType<RtfGeneratedText>());
        generated.FallbackText = "GENERATED";
        using (PdfCore.PdfDocument model = document.ToPdf(PortableOptions())) {
            var empty = Assert.IsType<PdfCore.RichParagraphBlock>(model.Blocks.ElementAt(1));
            PdfCore.PdfTextRun strut = Assert.Single(empty.Runs, run => run.IsLineBoxStrut);
            Assert.Equal(18D, strut.FontSize!.Value.ToPoints());
            Assert.Equal(PdfCore.PdfStandardFont.Courier, strut.StandardFont);
            using PdfPigDocument pdf = PdfPigDocument.Open(model.ToBytes());
            AssertPaint(pdf, "GENERATED", 18D, "Courier");
        }
        document.ApplyTheme(CustomTheme());
        Assert.Same(generated, Assert.Single(document.Paragraphs[0].Inlines.OfType<RtfGeneratedText>()));
        Assert.Equal(RtfGeneratedTextKind.PageNumber, generated.Kind);
        using PdfCore.PdfDocument themed = document.ToPdf(PortableOptions());
        PdfCore.PdfTextRun themedStrut = Assert.Single(Assert.IsType<PdfCore.RichParagraphBlock>(themed.Blocks.ElementAt(1)).Runs,
            run => run.IsLineBoxStrut);
        Assert.Equal(13D, themedStrut.FontSize!.Value.ToPoints());
        Assert.Equal(PdfCore.PdfStandardFont.Courier, themedStrut.StandardFont);
        using PdfPigDocument themedPdf = PdfPigDocument.Open(themed.ToBytes());
        AssertPaint(themedPdf, "GENERATED", 13D, "Courier");
        using RtfDocument reopened = RtfDocument.ParseResult(document.ToRtf()).Document;
        int? inheritedFont = reopened.Styles.Single(style => style.Id == 1 && style.Kind == RtfStyleKind.Paragraph).FontId;
        Assert.Equal("Courier New", reopened.Fonts.Single(font => font.Id == inheritedFont).Name);
    }

    private static RtfRun Run(RtfDocument document, string marker) => Assert.Single(document.Paragraphs.SelectMany(paragraph => paragraph.Runs),
        run => run.Text.Contains(marker, StringComparison.Ordinal));
    private static string Font(RtfDocument document, RtfRun run) => document.Fonts.Single(font => font.Id == run.FontId).Name;
    private static void AssertPaint(PdfPigDocument pdf, string marker, double size, string font) {
        var word = Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()), word => word.Text == marker);
        Assert.All(word.Letters, letter => {
            Assert.Equal(size, letter.PointSize, 2);
            Assert.Contains(font, letter.FontName, StringComparison.OrdinalIgnoreCase);
        });
    }
    private static RtfPdfExportOptions PortableOptions() => new() {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(), TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
    };
    private static OfficeTheme CustomTheme() => OfficeTheme.Classic
        .WithFonts(new OfficeThemeFonts(new OfficeThemeFont("Times New Roman", new[] { "serif" }), new OfficeThemeFont("Arial", new[] { "sans-serif" })))
        .WithTypeScale(OfficeTheme.Classic.TypeScale.With(body: OfficeLength.Points(13), heading1: OfficeLength.Points(23), heading6: OfficeLength.Points(11)))
        .WithSpacing(OfficeTheme.Classic.Spacing.With(paragraphAfter: OfficeLength.Points(7)))
        .WithTextStyle(OfficeTheme.Classic.TextStyle.With(headingBold: true, headingColor: OfficeColor.Red))
        .WithTableStyle(OfficeTheme.Classic.TableStyle.With(headerFill: OfficeColor.Blue, headerText: OfficeColor.White,
            cellPadding: OfficeLength.Points(4), banding: false));

    private const string Source = @"{\rtf1\ansi\deff0
{\fonttbl{\f0\fswiss Arial;}{\f1\fmodern Courier New;}{\f2\ftech\fcharset2 Symbol;}}
{\stylesheet{\s0\f0\fs18 Normal;}{\s1\sbasedon0\outlinelevel0 Heading;}{\s2\sbasedon0\outlinelevel8 Deep;}{\*\cs3\additive\i Emphasis;}}
\paperw12240\paperh15840\margl900\margr1000\margt1100\margb1200
\pard\s1 HEADING\par
\pard\s2 DEEPEST\par
\pard\s0\ltrpar\lang1033\fs18 PLAIN {\b BOLD} {\cs3 ITALIC} {\f1 CODE} {\f2 SYMBOL}
{\field{\*\fldinst HYPERLINK ""https://example.com/""}{\fldrslt LINK}}
\chftn{\footnote\pard\s0 NOTEBODY}\par
\pard\s0{\listtext\f2 S\tab} LISTBODY\par
\pard\s0{\pict\pngblip\picw1\pich1\picwgoal180\pichgoal240 PICTURE}\par
\trowd\trhdr\trkeep\cellx2500\cellx5000\pard\intbl\s0 HEADERA\cell HEADERB\cell\row
\trowd\trrh-600\trqr\clmgf\cellx2500\clmrg\cellx5000\pard\intbl\s0\qr VALUE\cell\cell\row
\pard\s0 SEPARATOR\par
\trowd\cellx2500\pard\intbl\s0 UNMARKED\cell\row
\pard\s0 END\par}";
}
