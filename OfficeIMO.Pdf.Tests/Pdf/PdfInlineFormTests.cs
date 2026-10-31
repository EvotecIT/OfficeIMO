using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInlineFormTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FillAndFlattenReplaceInlineCachedGlyphsWithoutGhosting(bool choice) {
        const string cached = "OBSOLETEVALUE";
        PdfInlineFormField descriptor = CreateDescriptor("Entry", choice, "OLD", cached);
        byte[] source = CreateParagraph(descriptor, cached, pageHeight: 160).ToBytes();
        PdfFormField initial = Assert.Single(PdfInspector.Inspect(source).FormFields);
        PdfFormWidget widget = Assert.Single(initial.Widgets);
        Assert.Equal("OLD", initial.Value);

        byte[] initialPaint = PdfDocument.Load(source).Forms.Flatten().ToBytes();
        Assert.Contains(cached, VisibleText(initialPaint));
        Assert.True(OldGlyphRegionInk(initialPaint, widget) > 0,
            "The sampled widget region must contain the original glyphs before filling.");

        PdfDocument filled = PdfDocument.Load(source).Forms.Fill(new Dictionary<string, string> { ["Entry"] = "NEW" });
        PdfFormField updated = Assert.Single(filled.Inspect().FormFields);
        Assert.Equal("NEW", updated.Value);
        Assert.Equal(WidgetGeometry(initial), WidgetGeometry(updated));
        Assert.DoesNotContain("OBSOLETE", VisibleText(filled.ToBytes()));

        byte[] flattened = filled.Forms.Flatten().ToBytes();
        Assert.Empty(PdfInspector.Inspect(flattened).FormFields);
        string visible = VisibleText(flattened);
        Assert.Contains(choice ? "OK" : "NEW", visible);
        Assert.DoesNotContain("OBSOLETE", visible);
        Assert.Contains("LEAD", visible);
        Assert.Contains("TAIL", visible);
        Assert.Equal(0, OldGlyphRegionInk(flattened, widget));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void WrappedInlineFieldsKeepOneLogicalValueAcrossLineWidgetsAndPages(bool choice, bool acrossPages) {
        string[] tokens = Enumerable.Range(1, 20).Select(index => "OBSOLETE" + index.ToString("D2")).ToArray();
        string cached = string.Join(" ", tokens);
        PdfInlineFormField descriptor = CreateDescriptor("Wrapped", choice, "OLD", cached);
        byte[] source = CreateParagraph(descriptor, cached, acrossPages ? 110 : 360, pageWidth: 200).ToBytes();
        PdfDocumentInfo initialInfo = PdfInspector.Inspect(source);
        PdfFormField initial = Assert.Single(initialInfo.FormFields);
        Assert.Equal("Wrapped", initial.Name);
        Assert.Equal("OLD", initial.Value);
        Assert.True(initial.WidgetCount > 1);
        Assert.All(initial.Widgets, widget => {
            Assert.Equal("Wrapped", widget.FieldName);
            Assert.True(widget.X2 > widget.X1);
            Assert.True(widget.Y2 > widget.Y1);
        });
        if (acrossPages) {
            Assert.True(initialInfo.PageCount > 1);
            Assert.True(initial.PageNumberCount > 1);
        } else {
            Assert.Equal(1, initialInfo.PageCount);
            Assert.Equal(new[] { 1 }, initial.PageNumbers);
        }
        string initialVisible = VisibleText(PdfDocument.Load(source).Forms.Flatten().ToBytes());
        foreach (string token in tokens) Assert.Contains(token, initialVisible);

        PdfDocument filled = PdfDocument.Load(source).Forms.Fill(new Dictionary<string, string> { ["Wrapped"] = "NEW" });
        PdfDocumentInfo filledInfo = filled.Inspect();
        PdfFormField updated = Assert.Single(filledInfo.FormFields);
        Assert.Equal("Wrapped", updated.Name);
        Assert.Equal("NEW", updated.Value);
        Assert.Equal(initialInfo.PageCount, filledInfo.PageCount);
        Assert.Equal(WidgetGeometry(initial), WidgetGeometry(updated));
        Assert.DoesNotContain("OBSOLETE", VisibleText(filled.ToBytes()));

        byte[] flattened = filled.Forms.Flatten().ToBytes();
        PdfDocumentInfo flattenedInfo = PdfInspector.Inspect(flattened);
        Assert.Empty(flattenedInfo.FormFields);
        Assert.Equal(initialInfo.PageCount, flattenedInfo.PageCount);
        string flattenedVisible = VisibleText(flattened);
        Assert.Contains(choice ? "OK" : "NEW", flattenedVisible);
        Assert.DoesNotContain("OBSOLETE", flattenedVisible);
        Assert.Contains("LEAD", flattenedVisible);
        Assert.Contains("TAIL", flattenedVisible);
    }

    [Fact]
    public void ReusedTableTextKeepsDistinctInlineFieldIdentitiesAndPlainText() {
        const string cached = "SAMECACHEDTEXT";
        PdfInlineFormField first = CreateDescriptor("First", false, "FIRST", cached);
        PdfInlineFormField second = CreateDescriptor("Second", false, "SECOND", cached);
        var rows = new[] { new[] {
            new PdfTableCell(new[] { PdfTextRun.Normal(cached).WithFormField(first) }),
            new PdfTableCell(new[] { PdfTextRun.Normal(cached).WithFormField(second) }),
            new PdfTableCell(new[] { PdfTextRun.Normal(cached) })
        } };
        byte[] source = PdfDocument.Create(Options(400, 160)).Table(rows, style: new PdfTableStyle {
            HeaderRowCount = 0, HeaderBold = false, FontSize = global::OfficeIMO.OfficeLength.Points(12), LineHeight = 1.2D,
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0), BorderColor = null,
            ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(110), global::OfficeIMO.OfficeLength.Points(110), global::OfficeIMO.OfficeLength.Points(110) }, MaxWidth = global::OfficeIMO.OfficeLength.Points(330), PreserveWidth = true
        }).ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(source);
        Assert.Equal(2, info.FormFields.Count);
        Assert.Equal("FIRST", info.FormFieldsByName["First"].Value);
        Assert.Equal("SECOND", info.FormFieldsByName["Second"].Value);
        Assert.Single(info.FormFieldsByName["First"].Widgets);
        Assert.Single(info.FormFieldsByName["Second"].Widgets);

        PdfDocument filled = PdfDocument.Load(source).Forms.Fill(new Dictionary<string, string> {
            ["First"] = "ONE", ["Second"] = "TWO"
        });
        Assert.Equal("ONE", filled.Inspect().FormFieldsByName["First"].Value);
        Assert.Equal("TWO", filled.Inspect().FormFieldsByName["Second"].Value);
        string visible = VisibleText(filled.Forms.Flatten().ToBytes());
        Assert.Contains("ONE", visible);
        Assert.Contains("TWO", visible);
        Assert.Equal(1, visible.Split(new[] { cached }, StringSplitOptions.None).Length - 1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnrelatedDuplicateNamesRemainRejected(bool inline) {
        PdfDocument document = PdfDocument.Create();
        if (inline) {
            PdfInlineFormField first = CreateDescriptor("Duplicate", false, "ONE", "ONE");
            PdfInlineFormField second = CreateDescriptor("Duplicate", false, "TWO", "TWO");
            document.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("ONE ").WithFormField(first), PdfTextRun.Normal("TWO").WithFormField(second)
            }));
        } else {
            document.TextField("Duplicate", value: "ONE").TextField("Duplicate", value: "TWO");
        }
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    private static PdfInlineFormField CreateDescriptor(string name, bool choice, string value, string cached) {
        var style = new PdfFormFieldStyle {
            BackgroundColor = null, BorderColor = null, BorderWidth = global::OfficeIMO.OfficeLength.Points(0), TextColor = global::OfficeIMO.OfficeColor.Black
        };
        PdfCanvasFormFieldItem item = choice
            ? PdfCanvasFormFieldItem.Choice(name, new[] {
                new PdfFormFieldOption(value, cached), new PdfFormFieldOption("NEW", "OK")
            }, new[] { value }, 0, 0, 1, 12, 12, true, false, style)
            : PdfCanvasFormFieldItem.Text(name, value, cached, 0, 0, 1, 12, 12, style, style);
        return new PdfInlineFormField(item);
    }

    private static PdfDocument CreateParagraph(PdfInlineFormField descriptor, string cached, double pageHeight,
        double pageWidth = 300) {
        int split = cached.IndexOf(' ', cached.Length / 2);
        if (split < 0) split = cached.Length / 2;
        else split++;
        // Styled runs share one descriptor, just as one Word SDT can contain
        // multiple runs and wrap across independently editable line widgets.
        return PdfDocument.Create(Options(pageWidth, pageHeight)).Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("LEAD "),
            PdfTextRun.Normal(cached.Substring(0, split)).WithFormField(descriptor),
            new PdfTextRun(cached.Substring(split), bold: true).WithFormField(descriptor),
            PdfTextRun.Normal(" TAIL")
        }), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), LineHeight = 1.2D });
    }

    private static PdfCreateOptions Options(double width, double height) => new() {
        PageWidth = global::OfficeIMO.OfficeLength.Points(width), PageHeight = global::OfficeIMO.OfficeLength.Points(height), MarginLeft = global::OfficeIMO.OfficeLength.Points(18), MarginRight = global::OfficeIMO.OfficeLength.Points(18),
        MarginTop = global::OfficeIMO.OfficeLength.Points(18), MarginBottom = global::OfficeIMO.OfficeLength.Points(18), DefaultFont = PdfStandardFont.Courier, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12)
    };

    private static (int? Page, double X1, double Y1, double X2, double Y2)[] WidgetGeometry(PdfFormField field) =>
        field.Widgets.OrderBy(widget => widget.PageNumber).ThenBy(widget => widget.Y1).ThenBy(widget => widget.X1)
            .Select(widget => (widget.PageNumber, widget.X1, widget.Y1, widget.X2, widget.Y2)).ToArray();

    private static string VisibleText(byte[] bytes) {
        using var pdf = PdfPigDocument.Open(bytes);
        return string.Join("\n", pdf.GetPages().Select(page => page.Text));
    }

    private static int OldGlyphRegionInk(byte[] flattened, PdfFormWidget widget) {
        Assert.NotNull(widget.PageNumber);
        using var pdf = PdfPigDocument.Open(flattened);
        double pageHeight = pdf.GetPage(widget.PageNumber!.Value).Height;
        PdfPageRenderResult rendered = Assert.Single(PdfDocument.Load(flattened).Render.Pages(
            PdfPageSelection.From(widget.PageNumber.Value), new PdfPageRenderOptions { Dpi = 72 }));
        Assert.True(OfficeRasterImageDecoder.TryDecode(rendered.Bytes!, out OfficeRasterImage? image));
        int ink = 0;
        for (int y = Math.Max(0, (int)Math.Floor(pageHeight - widget.Y2));
            y < Math.Min(image!.Height, (int)Math.Ceiling(pageHeight - widget.Y1)); y++) {
            for (int x = Math.Max(0, (int)Math.Floor(widget.X2 - 14D));
                x < Math.Min(image.Width, (int)Math.Ceiling(widget.X2 - 1D)); x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.R < 100 && pixel.G < 100 && pixel.B < 100) ink++;
            }
        }
        return ink;
    }
}
