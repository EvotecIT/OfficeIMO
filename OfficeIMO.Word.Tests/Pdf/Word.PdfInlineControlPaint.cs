using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class WordPdfInlineContentControlTests {
    [Theory]
    [InlineData("top")]
    [InlineData("center")]
    [InlineData("bottom")]
    public void InlineCheckBoxFontLineAlignmentMatchesItsOrdinaryCachedRun(string alignment) {
        using WordDocument source = CreateDocument(false, out WordParagraph paragraph);
        using WordDocument reference = CreateDocument(false, out WordParagraph referenceParagraph);
        foreach (WordParagraph target in new[] { paragraph, referenceParagraph }) {
            target.LineSpacingPoints = 30;
            target.LineSpacingRule = WordLineSpacingRule.Exact;
            target._paragraph.ParagraphProperties!.TextAlignment = new W.TextAlignment {
                Val = new VerticalTextAlignmentValues(alignment)
            };
            AddText(target, "LEAD ", 10);
        }

        WordCheckBox control = paragraph.AddCheckBox(true, tag: "AlignedCheck");
        W14.SdtContentCheckBox properties = control._sdtRun.SdtProperties!.GetFirstChild<W14.SdtContentCheckBox>()!;
        properties.GetFirstChild<W14.CheckedState>()!.Val = "0048";
        properties.GetFirstChild<W14.CheckedState>()!.Font = "Arial";
        properties.GetFirstChild<W14.UncheckedState>()!.Val = "004F";
        properties.GetFirstChild<W14.UncheckedState>()!.Font = "Arial";
        W.Run cache = control._sdtRun.SdtContentRun!.GetFirstChild<W.Run>()!;
        cache.RunProperties = new W.RunProperties(new W.RunFonts {
            Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial"
        }, new W.FontSize { Val = "20" });
        cache.GetFirstChild<W.Text>()!.Text = "H";
        referenceParagraph._paragraph.Append(cache.CloneNode(true));
        AddText(paragraph, " TAIL", 10);
        AddText(referenceParagraph, " TAIL", 10);

        byte[] sourceBytes = ConvertSavedDocument(source, InlinePaintOptions());
        Assert.True(Assert.Single(PdfCore.PdfReadDocument.Open(sourceBytes).FormFields).IsCheckBox);
        using PdfPigDocument actual = OpenVisiblePdf(sourceBytes);
        using PdfPigDocument expected = PdfPigDocument.Open(ConvertSavedDocument(reference, InlinePaintOptions()));
        var actualPage = actual.GetPage(1);
        var expectedPage = expected.GetPage(1);
        foreach (string token in new[] { "H", "TAIL" }) {
            var actualWord = Assert.Single(actualPage.GetWords(), word => word.Text == token);
            var expectedWord = Assert.Single(expectedPage.GetWords(), word => word.Text == token);
            Assert.Equal(ResolvedInlineControlFontName(expectedWord.Letters[0].FontName),
                ResolvedInlineControlFontName(actualWord.Letters[0].FontName));
            Assert.Equal(10D, actualWord.Letters[0].FontSize, 2);
            Assert.InRange(Math.Abs(actualWord.Letters[0].StartBaseLine.X - expectedWord.Letters[0].StartBaseLine.X), 0D, .05D);
            Assert.InRange(Math.Abs(actualWord.Letters[0].StartBaseLine.Y - expectedWord.Letters[0].StartBaseLine.Y), 0D, .05D);
        }
        var glyph = Assert.Single(actualPage.GetWords(), word => word.Text == "H");
        var tail = Assert.Single(actualPage.GetWords(), word => word.Text == "TAIL");
        Assert.InRange(Math.Abs(glyph.Letters[0].StartBaseLine.Y - tail.Letters[0].StartBaseLine.Y), 0D, .05D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InlineDateHighlightKeepsFullOrdinaryPaintAsAReportedStaticControl(bool inTable) {
        const string cached = "HIGHLIGHT";
        using WordDocument source = CreateDocument(inTable, out WordParagraph paragraph);
        using WordDocument reference = CreateDocument(inTable, out WordParagraph referenceParagraph);
        foreach (WordParagraph target in new[] { paragraph, referenceParagraph }) {
            target._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphMarkRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "20" }));
            AddText(target, "LEAD", 10);
        }
        WordDatePicker date = paragraph.AddDatePicker(new DateTime(2026, 5, 29), tag: "HighlightedDate");
        var cache = new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
            new W.FontSize { Val = "20" }, new W.Highlight { Val = W.HighlightColorValues.Yellow }),
            new W.Text(" " + cached) { Space = SpaceProcessingModeValues.Preserve });
        date._sdtRun.SdtContentRun = new W.SdtContentRun(cache);
        referenceParagraph._paragraph.Append(cache.CloneNode(true));
        AddText(paragraph, " TAIL", 10);
        AddText(referenceParagraph, " TAIL", 10);

        // Including the separator in the highlighted run makes its padded paint
        // overlap LEAD. The ordinary paragraph paints that background below all glyphs.
        string controlBefore = date._sdtRun.OuterXml;
        var converted = ConvertSavedInlinePaintWithReport(source);
        Assert.Equal(controlBefore, date._sdtRun.OuterXml);
        byte[] sourceBytes = converted.Bytes;
        byte[] referenceBytes = ConvertSavedDocument(reference, InlinePaintOptions());
        Assert.Empty(PdfCore.PdfReadDocument.Open(sourceBytes).FormFields);
        Assert.Single(converted.Result.Warnings, warning => warning.Code == "native-inline-control-background-static");
        var expectedPaint = InlineControlYellowBounds(referenceBytes);
        var actualPaint = InlineControlYellowBounds(sourceBytes);
        Assert.True(expectedPaint.Pixels > 20, "The ordinary highlighted run must provide a visible paint control.");
        using (PdfPigDocument expected = PdfPigDocument.Open(referenceBytes)) {
            var word = Assert.Single(expected.GetPage(1).GetWords(), word => word.Text == cached);
            Assert.True(expectedPaint.Left < Math.Floor(word.BoundingBox.Left * 2D));
            Assert.True(expectedPaint.Right > Math.Ceiling(word.BoundingBox.Right * 2D));
        }
        Assert.Equal(expectedPaint, actualPaint);
        AssertInlineControlPaintMatchesOrdinary(sourceBytes, referenceBytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InlineControlWithLaterInheritedBackgroundKeepsTheWholeControlStatic(bool inTable) {
        const string cached = "PLAIN HIGHLIGHT";
        using WordDocument source = CreateDocument(inTable, out WordParagraph paragraph);
        using WordDocument reference = CreateDocument(inTable, out WordParagraph referenceParagraph);
        foreach (WordDocument document in new[] { source, reference }) {
            W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new W.Style(new W.StyleName { Val = "Inline background base" },
                new W.StyleRunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "FFFF00" })) {
                Type = W.StyleValues.Character, StyleId = "InlineBackgroundBase"
            });
            styles.Append(new W.Style(new W.StyleName { Val = "Inline background inherited" },
                new W.BasedOn { Val = "InlineBackgroundBase" }) {
                Type = W.StyleValues.Character, StyleId = "InlineBackgroundInherited"
            });
        }
        foreach (WordParagraph target in new[] { paragraph, referenceParagraph }) {
            target._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphMarkRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "20" }));
            AddText(target, "LEAD ", 10);
        }
        W.SdtRun control = paragraph.AddComboBox(new[] { cached, "Other option" },
            tag: "LaterBackgroundRun", defaultValue: cached)._sdtRun;
        var plain = new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
            new W.FontSize { Val = "20" }), new W.Text("PLAIN"));
        var highlighted = new W.Run(new W.RunProperties(new W.RunStyle { Val = "InlineBackgroundInherited" },
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "20" }),
            new W.Text(" HIGHLIGHT") { Space = SpaceProcessingModeValues.Preserve });
        control.SdtContentRun = new W.SdtContentRun(plain, highlighted);
        referenceParagraph._paragraph.Append(plain.CloneNode(true), highlighted.CloneNode(true));
        AddText(paragraph, " TAIL", 10);
        AddText(referenceParagraph, " TAIL", 10);

        string controlBefore = control.OuterXml;
        var converted = ConvertSavedInlinePaintWithReport(source);
        Assert.Equal(controlBefore, control.OuterXml);
        Assert.Empty(PdfCore.PdfReadDocument.Open(converted.Bytes).FormFields);
        Assert.Single(converted.Result.Warnings, warning => warning.Code == "native-inline-control-background-static");
        byte[] referenceBytes = ConvertSavedDocument(reference, InlinePaintOptions());
        Assert.True(InlineControlYellowBounds(referenceBytes).Pixels > 20,
            "The later cached run must inherit visible background through its character style.");
        Assert.Equal(InlineControlYellowBounds(referenceBytes), InlineControlYellowBounds(converted.Bytes));
        AssertInlineControlPaintMatchesOrdinary(converted.Bytes, referenceBytes);
    }

    private static (byte[] Bytes, PdfCore.PdfDocumentConversionResult Result) ConvertSavedInlinePaintWithReport(WordDocument source) {
        byte[] docx = source.ToBytes();
        using (WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(docx), false)) {
            ValidationErrorInfo[] errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(package).ToArray();
            Assert.True(errors.Length == 0,
                string.Join("\n", errors.Select(error => error.Description + " @ " + error.Path?.XPath)));
        }
        using WordDocument loaded = WordDocument.Load(new MemoryStream(docx));
        string documentBefore = loaded._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        string stylesBefore = loaded._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart!.Styles!.OuterXml;
        PdfCore.PdfDocumentConversionResult result = loaded.ConvertToPdfWithEvidence(InlinePaintOptions());
        byte[] bytes = result.ToBytes();
        Assert.Equal(documentBefore, loaded._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Equal(stylesBefore, loaded._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.OuterXml);
        return (bytes, result);
    }

    private static void AssertInlineControlPaintMatchesOrdinary(byte[] actualBytes, byte[] expectedBytes) {
        using PdfPigDocument actual = PdfPigDocument.Open(actualBytes);
        using PdfPigDocument expected = PdfPigDocument.Open(expectedBytes);
        var actualPage = Assert.Single(actual.GetPages());
        var expectedPage = Assert.Single(expected.GetPages());
        Assert.Equal(expectedPage.Text, actualPage.Text);
        Assert.Equal(expectedPage.Letters.Count, actualPage.Letters.Count);
        for (int index = 0; index < expectedPage.Letters.Count; index++) {
            var expectedLetter = expectedPage.Letters[index];
            var actualLetter = actualPage.Letters[index];
            Assert.Equal(expectedLetter.Value, actualLetter.Value);
            Assert.Equal(ResolvedInlineControlFontName(expectedLetter.FontName), ResolvedInlineControlFontName(actualLetter.FontName));
            Assert.Equal(expectedLetter.FontSize, actualLetter.FontSize, 2);
            Assert.InRange(Math.Abs(expectedLetter.StartBaseLine.X - actualLetter.StartBaseLine.X), 0D, .05D);
            Assert.InRange(Math.Abs(expectedLetter.StartBaseLine.Y - actualLetter.StartBaseLine.Y), 0D, .05D);
        }
        var actualRaster = InlineControlPaintRaster(actualBytes);
        var expectedRaster = InlineControlPaintRaster(expectedBytes);
        Assert.Equal(expectedRaster.Width, actualRaster.Width);
        Assert.Equal(expectedRaster.Height, actualRaster.Height);
        Assert.Equal(expectedRaster.Pixels, actualRaster.Pixels);
    }

    private static (int Width, int Height, byte[] Pixels) InlineControlPaintRaster(byte[] bytes) {
        PdfCore.PdfPageRenderResult rendered = Assert.Single(PdfCore.PdfDocument.Load(bytes).Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Dpi = 144 }));
        Assert.True(rendered.Succeeded, string.Join("\n", rendered.Diagnostics));
        Assert.True(OfficeRasterImageDecoder.TryDecode(rendered.Bytes!, out OfficeRasterImage? image));
        OfficeRasterImage bitmap = image!;
        var pixels = new byte[bitmap.Width * bitmap.Height * 4];
        for (int y = 0; y < bitmap.Height; y++) for (int x = 0; x < bitmap.Width; x++) {
            OfficeColor color = bitmap.GetPixel(x, y);
            int offset = (y * bitmap.Width + x) * 4;
            pixels[offset] = color.R; pixels[offset + 1] = color.G;
            pixels[offset + 2] = color.B; pixels[offset + 3] = color.A;
        }
        return (bitmap.Width, bitmap.Height, pixels);
    }

    private static WordPdfExportOptions InlinePaintOptions() => new() {
        IncludePageNumbers = false, PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(160)),
        Margins = new PdfCore.PdfPageMargins(OfficeIMO.OfficeLength.Points(24),OfficeIMO.OfficeLength.Points(24),OfficeIMO.OfficeLength.Points(24),OfficeIMO.OfficeLength.Points(24))
    };

    private static (int Left, int Top, int Right, int Bottom, int Pixels) InlineControlYellowBounds(byte[] bytes) {
        PdfCore.PdfPageRenderResult rendered = Assert.Single(PdfCore.PdfDocument.Load(bytes).Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Dpi = 144 }));
        Assert.True(rendered.Succeeded, string.Join("\n", rendered.Diagnostics));
        Assert.True(OfficeRasterImageDecoder.TryDecode(rendered.Bytes!, out OfficeRasterImage? image));
        int left = int.MaxValue, top = int.MaxValue, right = -1, bottom = -1, pixels = 0;
        for (int y = 0; y < image!.Height; y++) for (int x = 0; x < image.Width; x++) {
            OfficeColor color = image.GetPixel(x, y);
            if (color.R < 220 || color.G < 220 || color.B > 40) continue;
            left = Math.Min(left, x); top = Math.Min(top, y);
            right = Math.Max(right, x); bottom = Math.Max(bottom, y); pixels++;
        }
        return (left, top, right, bottom, pixels);
    }
}
