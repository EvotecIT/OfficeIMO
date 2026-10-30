using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInlineFormPaintTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WidgetOnlyNamedFontIsDeclaredForEveryPageAndAppearanceTextOperator(bool precedingSpace) {
        const string family = "Inline widget resource control";
        var options = new PdfCreateOptions { DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family,
            ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics()));
        var checkedRun = PdfTextRun.Normal("A", fontSize: global::OfficeIMO.OfficeLength.Points(10), fontFamily: family);
        var uncheckedRun = PdfTextRun.Normal("B", fontSize: global::OfficeIMO.OfficeLength.Points(10), fontFamily: family);
        var field = new PdfInlineFormField(PdfCanvasFormFieldItem.CheckBox("OnlyWidgetFont", true,
            "Yes", "Yes", 0, 0, 12, 12, new PdfFormFieldStyle {
                BackgroundColor = null, BorderColor = null, BorderWidth = 0
            }), checkedRun, uncheckedRun);
        var widget = new PdfInlineFormWidget(field, 12, 12, sourceRun: checkedRun);
        byte[] bytes = PdfDocument.Create(options).Compose(compose => compose.Page(page => page.Content(content =>
            content.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal(precedingSpace ? "LEAD " : "LEAD"), PdfTextRun.Inline(widget), PdfTextRun.Normal(" TAIL")
            }))))).ToBytes();

        Assert.True(Assert.Single(PdfReadDocument.Open(bytes).FormFields).IsCheckBox);
        var (objects, _) = PdfSyntax.ParseObjects(bytes);
        PdfObject Resolve(PdfObject value) => value is PdfReference reference
            ? objects[reference.ObjectNumber].Value : value;
        PdfDictionary Dictionary(PdfObject value) => Assert.IsType<PdfDictionary>(Resolve(value));
        string StreamText(PdfStream stream) => Encoding.ASCII.GetString(
            OfficeIMO.Pdf.Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));

        PdfDictionary pageDictionary = Assert.Single(objects.Values.Select(value => value.Value)
            .OfType<PdfDictionary>(), value => value.Get<PdfName>("Type")?.Name == "Page");
        PdfDictionary pageFonts = Dictionary(Dictionary(pageDictionary.Items["Resources"]).Items["Font"]);
        var pageContent = Assert.IsType<PdfStream>(Resolve(pageDictionary.Items["Contents"]));
        string[] pageUses = TextFontNames(StreamText(pageContent));
        Assert.NotEmpty(pageUses);
        Assert.All(pageUses, name => Assert.True(pageFonts.Items.ContainsKey(name),
            "Page Tf references undeclared font /" + name));

        PdfDictionary annotation = Assert.Single(objects.Values.Select(value => value.Value)
            .OfType<PdfDictionary>(), value => value.Get<PdfName>("Subtype")?.Name == "Widget");
        PdfDictionary states = Dictionary(Dictionary(annotation.Items["AP"]).Items["N"]);
        Assert.True(states.Items.ContainsKey("Off"));
        Assert.True(states.Items.ContainsKey("Yes"));
        foreach (PdfObject state in states.Items.Values) {
            var appearance = Assert.IsType<PdfStream>(Resolve(state));
            PdfDictionary appearanceFonts = Dictionary(Dictionary(appearance.Dictionary.Items["Resources"]).Items["Font"]);
            string[] appearanceUses = TextFontNames(StreamText(appearance));
            Assert.NotEmpty(appearanceUses);
            Assert.All(appearanceUses, name => Assert.True(appearanceFonts.Items.ContainsKey(name),
                "Widget appearance Tf references undeclared font /" + name));
        }

        PdfPageRenderResult rendered = Assert.Single(PdfDocument.Load(bytes).Render.Pages(
            options: new PdfPageRenderOptions { Dpi = 72 }));
        Assert.True(rendered.Succeeded, string.Join("\n", rendered.Diagnostics));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.IndexOf("font", StringComparison.OrdinalIgnoreCase) >= 0 &&
            (diagnostic.IndexOf("missing", StringComparison.OrdinalIgnoreCase) >= 0 ||
             diagnostic.IndexOf("unknown", StringComparison.OrdinalIgnoreCase) >= 0 ||
             diagnostic.IndexOf("unresolved", StringComparison.OrdinalIgnoreCase) >= 0));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InlineControlBackgroundRequiresStaticParagraphPaint(bool checkBox) {
        var style = new PdfFormFieldStyle { BackgroundColor = null, BorderColor = null, BorderWidth = 0 };
        var highlighted = PdfTextRun.Normal("CACHE", fontSize: global::OfficeIMO.OfficeLength.Points(10),
            backgroundColor: global::OfficeIMO.OfficeColor.FromRgb(255, 255, 0));
        var field = new PdfInlineFormField(checkBox
            ? PdfCanvasFormFieldItem.CheckBox("Background", true, "Yes", "Yes", 0, 0, 10, 10, style)
            : PdfCanvasFormFieldItem.Text("Background", "CACHE", "CACHE", 0, 0, 1, 10, 10, style, style));
        Assert.Throws<NotSupportedException>(() => {
            if (checkBox) _ = new PdfInlineFormWidget(field, 10, 10, sourceRun: highlighted);
            else _ = highlighted.WithFormField(field);
        });
    }
    [Fact]
    public void InlineTextAppearanceRetainsItalicInkBeyondTheGlyphAdvance() {
        const string cached = "fff";
        var style = new PdfFormFieldStyle { BackgroundColor = null, BorderColor = null, BorderWidth = 0 };
        var field = new PdfInlineFormField(PdfCanvasFormFieldItem.Text("Italic", cached, cached,
            0, 0, 1, 16, 16, style, style));
        byte[] Create(bool inline) {
            PdfTextRun run = PdfTextRun.Normal(cached, fontSize: global::OfficeIMO.OfficeLength.Points(16), font: PdfStandardFont.TimesItalic);
            if (inline) run = run.WithFormField(field);
            return PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(300), PageHeight = global::OfficeIMO.OfficeLength.Points(120), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), DefaultFont = PdfStandardFont.TimesRoman, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(16)
            }).Compose(compose => compose.Page(page => page.Content(content => content.Paragraph(paragraph =>
                paragraph.Runs(new[] { PdfTextRun.Normal("LEAD "), run, PdfTextRun.Normal("       TAIL") }))))).ToBytes();
        }

        byte[] source = Create(true);
        byte[] visible = PdfDocument.Load(source).Forms.Flatten().ToBytes();
        byte[] reference = Create(false);
        using var actualPdf = PdfPigDocument.Open(visible);
        using var expectedPdf = PdfPigDocument.Open(reference);
        var actualWord = Assert.Single(actualPdf.GetPage(1).GetWords(), word => word.Text == cached);
        var expectedWord = Assert.Single(expectedPdf.GetPage(1).GetWords(), word => word.Text == cached);
        Assert.Equal(expectedWord.Letters[0].StartBaseLine.X, actualWord.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(expectedWord.Letters[0].StartBaseLine.Y, actualWord.Letters[0].StartBaseLine.Y, 3);

        var expected = InkRaster(reference);
        var actual = InkRaster(visible);
        Assert.Equal(expected.Width, actual.Width);
        Assert.Equal(expected.Height, actual.Height);
        int advanceEnd = (int)Math.Ceiling(expectedWord.Letters[expectedWord.Letters.Count - 1].EndBaseLine.X * 2D);
        var tail = Assert.Single(expectedPdf.GetPage(1).GetWords(), word => word.Text == "TAIL");
        int nextWordStart = (int)Math.Floor(tail.BoundingBox.Left * 2D);
        int referenceOverhang = 0, missingOverhang = 0;
        for (int y = 0; y < expected.Height; y++) for (int x = advanceEnd + 2; x < nextWordStart; x++) {
            if (!expected.Ink[y * expected.Width + x]) continue;
            referenceOverhang++;
            if (!HasInkWithinOnePixel(actual, x, y)) missingOverhang++;
        }
        Assert.True(referenceOverhang > 0, "The ordinary italic control must paint beyond its glyph advance.");
        Assert.True(missingOverhang == 0,
            missingOverhang + " ordinary italic ink pixels beyond the advance are missing from the inline appearance.");
        Assert.Equal(0, MissingInk(expected, actual));
        Assert.Equal(0, MissingInk(actual, expected));
    }

    [Fact]
    public void FillingMultiwordUnderlinedInlineTextRemovesCachedGlyphsAndSeparatorDecoration() {
        const string cached = "OBSOLETE VALUE";
        var style = new PdfFormFieldStyle { BackgroundColor = null, BorderColor = null, BorderWidth = 0 };
        var field = new PdfInlineFormField(PdfCanvasFormFieldItem.Text("Underlined", "OLD", cached,
            0, 0, 1, 10, 10, style, style));
        byte[] source = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(300), PageHeight = global::OfficeIMO.OfficeLength.Points(120), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
            MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        }).Compose(compose => compose.Page(page => page.Content(content => content.Paragraph(paragraph =>
            paragraph.Runs(new[] {
                PdfTextRun.Normal("LEAD "),
                new PdfTextRun(cached, underline: true, fontSize: global::OfficeIMO.OfficeLength.Points(10), font: PdfStandardFont.Helvetica)
                    .WithFormField(field),
                PdfTextRun.Normal(" TAIL")
            }))))).ToBytes();
        PdfFormWidget widget = Assert.Single(Assert.Single(PdfReadDocument.Open(source).FormFields).Widgets);
        byte[] initial = PdfDocument.Load(source).Forms.Flatten().ToBytes();
        using var initialPdf = PdfPigDocument.Open(initial);
        var initialWords = initialPdf.GetPage(1).GetWords().ToArray();
        var obsolete = Assert.Single(initialWords, word => word.Text == "OBSOLETE");
        var value = Assert.Single(initialWords, word => word.Text == "VALUE");
        int separatorLeft = (int)Math.Ceiling(obsolete.Letters.Last().EndBaseLine.X * 2D);
        int separatorRight = (int)Math.Floor(value.Letters[0].StartBaseLine.X * 2D);
        Assert.True(separatorRight > separatorLeft, "The cached words must have a visible separator region.");

        byte[] filled = PdfDocument.Load(source).Forms.Fill(new Dictionary<string, string> {
            ["Underlined"] = "NEW"
        }).ToBytes();
        Assert.Equal("NEW", Assert.Single(PdfReadDocument.Open(filled).FormFields).Value);
        byte[] flattened = PdfDocument.Load(filled).Forms.Flatten().ToBytes();
        Assert.Empty(PdfReadDocument.Open(flattened).FormFields);
        using var flattenedPdf = PdfPigDocument.Open(flattened);
        var words = flattenedPdf.GetPage(1).GetWords().ToArray();
        var replacement = Assert.Single(words, word => word.Text == "NEW");
        Assert.DoesNotContain(words, word => word.Text == "OBSOLETE" || word.Text == "VALUE");
        Assert.Contains(words, word => word.Text == "LEAD");
        Assert.Contains(words, word => word.Text == "TAIL");

        var initialInk = InkRaster(initial);
        var filledInk = InkRaster(flattened);
        int oldRegionLeft = (int)Math.Ceiling(replacement.Letters.Last().EndBaseLine.X * 2D) + 2;
        int oldRegionRight = (int)Math.Floor(widget.X2 * 2D) - 1;
        Assert.True(oldRegionLeft < separatorLeft && oldRegionRight > separatorRight,
            "The cached separator must lie beyond the shorter replacement's painted region.");
        Assert.True(InkBetweenX(initialInk, oldRegionLeft, oldRegionRight) > 0,
            "The old region must contain cached glyph ink before filling.");
        Assert.True(InkBetweenX(initialInk, separatorLeft, separatorRight) > 0,
            "The cached inter-word underline must provide visible separator ink before filling.");
        Assert.Equal(0, InkBetweenX(filledInk, oldRegionLeft, oldRegionRight));
        Assert.Equal(0, InkBetweenX(filledInk, separatorLeft, separatorRight));
    }

    [Fact]
    public void ToggledInlineCheckBoxFitsAlternateGlyphInkWithoutChangingItsFrameOrNeighbors() {
        const string family = "Toggle outline control";
        const double size = 20D, pageWidth = 300D, pageHeight = 120D;
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth), PageHeight = global::OfficeIMO.OfficeLength.Points(pageHeight), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
            MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        }.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 'B')));
        var checkedRun = PdfTextRun.Italicized("A", fontSize: global::OfficeIMO.OfficeLength.Points(size), fontFamily: family);
        var uncheckedRun = PdfTextRun.Italicized("B", fontSize: global::OfficeIMO.OfficeLength.Points(size), fontFamily: family);
        var metrics = Assert.IsType<PdfWriter.PositionedTextBoxMeasurement>(PdfWriter.MeasurePositionedTextBox(checkedRun, options));
        var alternateMetrics = Assert.IsType<PdfWriter.PositionedTextBoxMeasurement>(PdfWriter.MeasurePositionedTextBox(uncheckedRun, options));
        Assert.Equal(metrics.Width, alternateMetrics.Width, 3);
        var style = new PdfFormFieldStyle { BackgroundColor = null, BorderColor = null, BorderWidth = 0 };
        var field = new PdfInlineFormField(PdfCanvasFormFieldItem.CheckBox("InkToggle", true,
            "Yes", "Yes", 0, 0, metrics.Width, metrics.Ascent + metrics.Descent, style), checkedRun, uncheckedRun);
        var widget = new PdfInlineFormWidget(field, metrics.Width, metrics.Ascent + metrics.Descent,
            -metrics.Descent, sourceRun: checkedRun);
        byte[] source = PdfDocument.Create(options).Compose(compose => compose.Page(page => page.Content(content =>
            content.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("LEFT"), PdfTextRun.Inline(widget),
                PdfTextRun.Inline(new PdfInlineBox(20D, 0.01D, borderWidth: 0D)), PdfTextRun.Normal("RIGHT")
            }))))).ToBytes();
        PdfFormWidget beforeWidget = Assert.Single(Assert.Single(PdfReadDocument.Open(source).FormFields).Widgets);
        byte[] initial = PdfDocument.Load(source).Forms.Flatten().ToBytes();
        using var initialPdf = PdfPigDocument.Open(initial);
        var initialLetters = initialPdf.GetPage(1).Letters;
        var originalGlyph = Assert.Single(initialLetters, letter => letter.Value == "A");

        byte[] filled = PdfDocument.Load(source).Forms.Fill(new Dictionary<string, string> {
            ["InkToggle"] = "false"
        }).ToBytes();
        PdfFormField updatedField = Assert.Single(PdfReadDocument.Open(filled).FormFields);
        Assert.Equal("Off", updatedField.Value);
        PdfFormWidget afterWidget = Assert.Single(updatedField.Widgets);
        Assert.Equal((beforeWidget.PageNumber, beforeWidget.X1, beforeWidget.Y1, beforeWidget.X2, beforeWidget.Y2),
            (afterWidget.PageNumber, afterWidget.X1, afterWidget.Y1, afterWidget.X2, afterWidget.Y2));
        byte[] flattened = PdfDocument.Load(filled).Forms.Flatten().ToBytes();
        Assert.Empty(PdfReadDocument.Open(flattened).FormFields);
        using var flattenedPdf = PdfPigDocument.Open(flattened);
        var filledLetters = flattenedPdf.GetPage(1).Letters;
        var alternateGlyph = Assert.Single(filledLetters, letter => letter.Value == "B");
        Assert.DoesNotContain(filledLetters, letter => letter.Value == "A");
        Assert.Equal(originalGlyph.StartBaseLine.Y, alternateGlyph.StartBaseLine.Y, 3);
        var beforeNeighbors = initialLetters.Where(letter => letter.Value != "A").ToArray();
        var afterNeighbors = filledLetters.Where(letter => letter.Value != "B").ToArray();
        Assert.Equal(beforeNeighbors.Length, afterNeighbors.Length);
        for (int index = 0; index < beforeNeighbors.Length; index++) {
            Assert.Equal(beforeNeighbors[index].Value, afterNeighbors[index].Value);
            Assert.Equal(beforeNeighbors[index].StartBaseLine.X, afterNeighbors[index].StartBaseLine.X, 3);
            Assert.Equal(beforeNeighbors[index].StartBaseLine.Y, afterNeighbors[index].StartBaseLine.Y, 3);
        }

        // The existing fixture gives A/B the same 500-unit advance, while their
        // rectangular outlines are 400/600 units wide and 700 units tall. The
        // same authored italic style shears each whole outline. Compress all
        // four alternate corners into the current widget, retaining its baseline.
        double height = 700D * size / 1000D;
        double alternateWidth = 600D * size / 1000D;
        double shear = ContentStreamBuilder.SyntheticObliqueShear;
        double fullInkWidth = alternateWidth + shear * height;
        double available = beforeWidget.X2 - originalGlyph.StartBaseLine.X;
        double scaleX = Math.Min(1D, available / fullInkWidth);
        Assert.True(fullInkWidth > available + 1D, "The alternate glyph must require an ink-based fit despite its unchanged advance.");
        OfficePoint[] corners = new[] {
            new OfficePoint(0D, 0D), new OfficePoint(alternateWidth, 0D),
            new OfficePoint(alternateWidth, height), new OfficePoint(0D, height)
        };
        PdfTextSpan alternateSpan = Assert.Single(PdfReadDocument.Open(flattened).Pages[0].GetTextSpans(),
            span => span.Text == "B");
        Assert.True(alternateSpan.TextToPageTransform.HasValue,
            "The emitted alternate glyph must retain its complete PDF text-to-page transform.");
        Matrix2D transform = alternateSpan.TextToPageTransform!.Value;
        // PDF text matrices and widget coordinates use at least three decimal
        // places. Propagate one 0.001-unit rounding bound through the fit and
        // shear coefficients, plus the baseline, widget edge and AP placement.
        const double numberPrecision = 0.001D;
        foreach (OfficePoint corner in corners) {
            var actual = transform.Transform(corner.X, corner.Y);
            double shearedX = corner.X + shear * corner.Y;
            double expectedX = originalGlyph.StartBaseLine.X + scaleX * shearedX;
            double expectedY = originalGlyph.StartBaseLine.Y + corner.Y;
            double xTolerance = 3D * numberPrecision + numberPrecision * Math.Abs(shearedX) +
                (Math.Abs(scaleX) + numberPrecision) * numberPrecision * Math.Abs(corner.Y);
            double yTolerance = numberPrecision * (3D + Math.Abs(corner.Y));
            Assert.InRange(actual.X, expectedX - xTolerance, expectedX + xTolerance);
            Assert.InRange(actual.Y, expectedY - yTolerance, expectedY + yTolerance);
            Assert.InRange(actual.X, beforeWidget.X1 - xTolerance, beforeWidget.X2 + xTolerance);
            Assert.InRange(actual.Y, beforeWidget.Y1 - yTolerance, beforeWidget.Y2 + yTolerance);
        }
    }

    private static int InkBetweenX((int Width, int Height, bool[] Ink) image, int left, int right) {
        int count = 0;
        for (int y = 0; y < image.Height; y++)
            for (int x = Math.Max(0, left); x < Math.Min(image.Width, right); x++)
                if (image.Ink[y * image.Width + x]) count++;
        return count;
    }

    private static (int Width, int Height, bool[] Ink) InkRaster(byte[] bytes) {
        PdfPageRenderResult rendered = Assert.Single(PdfDocument.Load(bytes).Render.Pages(
            options: new PdfPageRenderOptions { Dpi = 144 }));
        Assert.True(rendered.Succeeded, string.Join("\n", rendered.Diagnostics));
        Assert.True(OfficeRasterImageDecoder.TryDecode(rendered.Bytes!, out OfficeRasterImage? image));
        OfficeRasterImage bitmap = image!;
        var ink = new bool[bitmap.Width * bitmap.Height];
        for (int y = 0; y < bitmap.Height; y++) for (int x = 0; x < bitmap.Width; x++) {
            OfficeColor color = bitmap.GetPixel(x, y);
            ink[y * bitmap.Width + x] = color.R < 100 && color.G < 100 && color.B < 100;
        }
        return (bitmap.Width, bitmap.Height, ink);
    }

    private static bool HasInkWithinOnePixel((int Width, int Height, bool[] Ink) image, int x, int y) {
        for (int dy = -1; dy <= 1; dy++) for (int dx = -1; dx <= 1; dx++) {
            int neighborX = x + dx, neighborY = y + dy;
            if (neighborX >= 0 && neighborX < image.Width && neighborY >= 0 && neighborY < image.Height &&
                image.Ink[neighborY * image.Width + neighborX]) return true;
        }
        return false;
    }

    private static int MissingInk((int Width, int Height, bool[] Ink) expected,
        (int Width, int Height, bool[] Ink) actual) {
        int missing = 0;
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++)
            if (expected.Ink[y * expected.Width + x] && !HasInkWithinOnePixel(actual, x, y)) missing++;
        return missing;
    }

    private static string[] TextFontNames(string content) => Regex.Matches(content,
        @"/(?<font>[^\s/<>\[\]()]+)\s+[-+]?(?:\d+\.?\d*|\.\d+)\s+Tf\b",
        RegexOptions.CultureInvariant).Cast<Match>().Select(match => match.Groups["font"].Value)
        .Distinct(StringComparer.Ordinal).ToArray();

}
