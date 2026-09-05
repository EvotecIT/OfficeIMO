using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReusableContentLayoutTests {
    [Fact]
    public void NoLossGuardRequiresSuccessfulOutput() {
        var failure = new IOException("Output unavailable");
        var result = PdfSaveResult.FromFailure("report.pdf", failure);
        var exception = Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Same(failure, exception.InnerException);
        var successful = PdfSaveResult.FromSuccess("report.pdf", 120);
        Assert.Same(successful, successful.RequireNoLoss());
    }

    private static PdfOptions Options() => new PdfOptions {
        PageWidth = 300, PageHeight = 220,
        MarginLeft = 20, MarginRight = 20, MarginTop = 20, MarginBottom = 20,
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 10
    };

    [Fact]
    public void IncrementalContentAndCallbackContentShareOneDocument() {
        var document = PdfDocument.Create(Options());
        document.Content.Text("Before");
        document.Content.Column(content => content.Text("Middle"));
        document.Content.Text("After");
        string text = PdfReadDocument.Open(document.ToBytes()).ExtractText();
        Assert.True(text.IndexOf("Before", StringComparison.Ordinal) < text.IndexOf("Middle", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Middle", StringComparison.Ordinal) < text.IndexOf("After", StringComparison.Ordinal));
    }

    [Fact]
    public void AutomaticWidthUsesWidestExplicitLine() {
        double Render(bool separateParagraphs) {
            var document = PdfDocument.Create(Options());
            document.Content.Row(row => row.Gap(10).AutoColumn(column => {
                if (separateParagraphs) {
                    for (int i = 0; i < 3; i++) column.Text("ABCDEFGHIJKLMNO");
                } else {
                    column.Paragraph(p => p.Text("ABCDEFGHIJKLMNO").LineBreak()
                        .Text("ABCDEFGHIJKLMNO").LineBreak().Text("ABCDEFGHIJKLMNO"));
                }
            }).RelativeColumn(column => column.Text("Relative")));
            using var parsed = PdfPigDocument.Open(document.ToBytes());
            return parsed.GetPage(1).GetWords().Single(word => word.Text == "Relative").BoundingBox.Left;
        }
        Assert.Equal(Render(true), Render(false), 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AutomaticWidthUsesRegisteredFamilyForStyledRuns(bool boldItalic) {
        const string familyName = "Auto Width Fixture";
        var options = Options();
        byte[] font = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont('W', ' ');
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(familyName, font, font, font, font));
        var document = PdfDocument.Create(options);
        document.Content.Row(row => row.Gap(10)
            .AutoColumn(column => column.Paragraph(p => p.FontFamily(familyName).Bold(boldItalic).Italic(boldItalic).Text("WWWW")))
            .RelativeColumn(column => column.Text("Relative")));
        using var parsed = PdfPigDocument.Open(document.ToBytes());
        // The fixture face has a 500-unit advance at 1000 units per em.
        Assert.Equal(50D, parsed.GetPage(1).GetWords().Single(word => word.Text == "Relative").BoundingBox.Left, 2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PercentageColumnsUseLiteralShareWithOrWithoutRelativeColumns(bool withRelative) {
        var document = PdfDocument.Create(Options());
        document.Content.Row(row => {
            row.Gap(0).PercentColumn(30, column => column.Text("First"))
                .PercentColumn(20, column => column.Text("Second"));
            if (withRelative) row.RelativeColumn(column => column.Text("Third"));
        });
        using var parsed = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(98D, parsed.GetPage(1).GetWords().Single(word => word.Text == "Second").BoundingBox.Left, 2);
    }

    [Fact]
    public void ComponentInDecoratedSemanticRowPreservesEveryPageAndCapture() {
        var capture = new PdfLayoutPositionCapture();
        var document = PdfDocument.Create(Options());
        document.Content.Row(row => row.Gap(10)
            .RelativeColumn(column => column.Panel(panel => panel.Semantic(PdfSemanticRole.Section,
                semantic => semantic.Component(new LinesComponent(), capture: capture)),
                new PdfPanelStyle { PaddingX = 8, PaddingY = 5, Background = PdfColor.FromRgb(240, 245, 250) }))
            .RelativeColumn(column => column.Text("Adjacent")));
        byte[] bytes = document.ToBytes();
        using var parsed = PdfPigDocument.Open(bytes);
        Assert.True(parsed.NumberOfPages > 1);
        string text = PdfReadDocument.Open(bytes).ExtractText();
        for (int index = 0; index < 30; index++) Assert.Contains("Line" + index.ToString("D2"), text);
        Assert.Equal(parsed.NumberOfPages, capture.Regions.Select(region => region.PageNumber).Distinct().Count());
        Assert.All(capture.Regions, region => {
            Assert.Equal(28D, region.X, 2);
            Assert.InRange(region.Y, 20D, 200D);
            Assert.True(region.Height > 0D);
        });
    }

    [Fact]
    public void RowPanelPreservesTableAndForm() {
        var document = PdfDocument.Create(Options());
        document.Content.Row(row => row.RelativeColumn(column => column.Panel(panel => panel
            .Table(new[] { new[] { "Name", "Value" }, new[] { "Item", "42" } })
            .TextField("answer"))));
        byte[] bytes = document.ToBytes();
        Assert.Contains("answer", PdfInspector.Inspect(bytes).FormFieldsByName.Keys);
        using var parsed = PdfPigDocument.Open(bytes);
        var words = parsed.GetPage(1).GetWords().ToArray();
        Assert.True(words.Single(word => word.Text == "Value").BoundingBox.Left > words.Single(word => word.Text == "Name").BoundingBox.Left + 30D);
    }

    [Fact]
    public void OversizedKeepTogetherComponentInRowFailsPromptly() {
        var document = PdfDocument.Create(Options());
        document.Content.Row(row => row.RelativeColumn(column => column.Component(new LinesComponent(),
            new PdfFlowOptions { KeepTogether = true })));
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    private sealed class LinesComponent : IPdfComponent {
        public void Compose(PdfContentBuilder content) {
            for (int index = 0; index < 30; index++) content.Text("Line" + index.ToString("D2"));
        }
    }

    [Fact]
    public void PaddedRowPanelRepeatsTableHeadersOnEveryContinuationPage() {
        var rows = new List<string[]> { new[] { "Header", "Value" } };
        for (int index = 0; index < 30; index++) rows.Add(new[] { "Entry" + index, "42" });
        var document = PdfDocument.Create(Options());
        document.Content.Row(row => row.RelativeColumn(column => column.Panel(panel =>
            panel.Table(rows, style: new PdfTableStyle { HeaderRowCount = 1, RepeatHeaderRowCount = 1 }),
            new PdfPanelStyle { PaddingY = 10 })));
        using var parsed = PdfPigDocument.Open(document.ToBytes());
        Assert.True(parsed.NumberOfPages > 1);
        foreach (var page in parsed.GetPages()) Assert.Contains("Header", page.Text);
    }

    [Fact]
    public void RowKeepTogetherIncludesFirstChildSpacingAfterPadding() {
        var document = PdfDocument.Create(Options());
        document.Content.Row(row => row.RelativeColumn(column => column.Panel(panel =>
            panel.Paragraph(p => {
                for (int i = 0; i < 6; i++) {
                    if (i > 0) p.LineBreak();
                    p.Text("Line");
                }
            }, style: new PdfParagraphStyle { SpacingBefore = 100, SpacingAfter = 0 }),
            new PdfPanelStyle { PaddingY = 10, KeepTogether = true })));
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }
}
