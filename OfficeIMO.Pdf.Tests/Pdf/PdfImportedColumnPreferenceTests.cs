using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfImportedColumnPreferenceTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    public void ImportedContentMinimumCompressionPreservesPaddingAndAllClippedCellText(int automaticColumn, bool explicitMaximum) {
        const string leftText = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        const string rightText = "ABCDEFGHIJKLM";
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(explicitMaximum ? 240D : 119D), PageHeight = global::OfficeIMO.OfficeLength.Points(300D),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(20D), MarginRight = global::OfficeIMO.OfficeLength.Points(20D), MarginTop = global::OfficeIMO.OfficeLength.Points(20D), MarginBottom = global::OfficeIMO.OfficeLength.Points(20D),
            DefaultFont = PdfStandardFont.Courier
        };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, AutoFitColumns = true, AutoFitWidthUsesContentMinimum = true,
            PreserveWidth = true, ClipTextToCellBounds = true,
            AllowWidthBeyondFrame = explicitMaximum, PreferredWidth = global::OfficeIMO.OfficeLength.Points(explicitMaximum ? 200D : 79D),
            MaxWidth = explicitMaximum ? global::OfficeIMO.OfficeLength.Points(79D) : null,
            AutoFitColumnPreferredWidthPoints = new List<double?> {
                automaticColumn == 1 ? null : 60D, automaticColumn == 2 ? null : 140D },
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(2D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D), FontSize = global::OfficeIMO.OfficeLength.Points(10D), BorderColor = null
        };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options)
            .Table(new[] { new[] { leftText, rightText } }, style: style).ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal(leftText + rightText, string.Concat(page.Letters.Select(letter => letter.Value)));
        Assert.InRange(page.Letters[leftText.Length].StartBaseLine.X - page.Letters[0].StartBaseLine.X,
            51.331D, 51.335D);
        Assert.All(page.Letters, letter => Assert.InRange(letter.BoundingBox.Right, 20D, 99.001D));
    }

    [Theory]
    [InlineData(100D, 31.428571D)]
    [InlineData(400D, 90D)]
    public void ImportedPreferredFrameRetainsAnExplicitMaximumBeyondTheTextFrame(double maximum, double expectedFirstWidth) {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(120D), PageHeight = global::OfficeIMO.OfficeLength.Points(200D),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(20D), MarginRight = global::OfficeIMO.OfficeLength.Points(20D), MarginTop = global::OfficeIMO.OfficeLength.Points(20D), MarginBottom = global::OfficeIMO.OfficeLength.Points(20D)
        };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, AutoFitColumns = true, AutoFitWidthUsesContentMinimum = true,
            AllowWidthBeyondFrame = true, PreferredWidth = global::OfficeIMO.OfficeLength.Points(300D), MaxWidth = global::OfficeIMO.OfficeLength.Points(maximum), PreserveWidth = true,
            AutoFitColumnPreferredWidthPoints = new List<double?> { 60D, 140D },
            ColumnMinWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(40D) },
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(0D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D), FontSize = global::OfficeIMO.OfficeLength.Points(10D), BorderColor = null
        };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options)
            .Table(new[] { new[] { "A", "B" } }, style: style).ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var left = Assert.Single(page.Letters, letter => letter.Value == "A");
        var right = Assert.Single(page.Letters, letter => letter.Value == "B");
        Assert.Equal(expectedFirstWidth, right.StartBaseLine.X - left.StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(100D, 300D, 0D, 0, 31.428571D)]
    [InlineData(100D, 300D, 25D, 0, 25D)]
    [InlineData(300D, 300D, 0D, 0, 90D)]
    [InlineData(300D, 100D, 0D, 0, 90D)]
    [InlineData(100D, 300D, 0D, 1, 20D)]
    [InlineData(300D, 300D, 0D, 1, 160D)]
    [InlineData(300D, 300D, 25D, 1, 25D)]
    [InlineData(100D, 300D, 0D, 2, 60D)]
    [InlineData(300D, 300D, 0D, 2, 60D)]
    public void ImportedPreferencesShrinkAfterMinimumsAndGrowProportionally(
        double preferredWidth, double frameWidth, double firstMaximum, int automaticColumn, double expectedFirstWidth) {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(frameWidth + 40D), PageHeight = global::OfficeIMO.OfficeLength.Points(200D),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(20D), MarginRight = global::OfficeIMO.OfficeLength.Points(20D), MarginTop = global::OfficeIMO.OfficeLength.Points(20D), MarginBottom = global::OfficeIMO.OfficeLength.Points(20D)
        };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, AutoFitColumns = true, AutoFitWidthUsesContentMinimum = true,
            AllowWidthBeyondFrame = true, PreferredWidth = global::OfficeIMO.OfficeLength.Points(preferredWidth), PreserveWidth = true,
            AutoFitColumnPreferredWidthPoints = new List<double?> {
                automaticColumn == 1 ? null : 60D, automaticColumn == 2 ? null : 140D },
            ColumnMinWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(40D) },
            ColumnMaxWidths = (firstMaximum > 0D ? new List<double?> { firstMaximum, null } : null) is { } lengths5269 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths5269, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null,
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(0D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D), FontSize = global::OfficeIMO.OfficeLength.Points(10D), BorderColor = null
        };
        byte[] bytes = PdfDocument.Create(options).Table(new[] { new[] { "A", "B" } }, style: style).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var left = Assert.Single(page.Letters, letter => letter.Value == "A");
        var right = Assert.Single(page.Letters, letter => letter.Value == "B");
        Assert.Equal(expectedFirstWidth, right.StartBaseLine.X - left.StartBaseLine.X, 3);
        Assert.Equal(left.StartBaseLine.Y, right.StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConstrainedImportedPreferencesSurviveSourceBudgetCompressionAndCapturedStyleChanges(bool oversizedWords) {
        string leftText = oversizedWords ? "ABCDEFGHIJKLMNOPQRSTUVWXYZ" : "A";
        string rightText = oversizedWords ? "ABCDEFGHIJKLM" : "ABCDEFGHIJK";
        PdfTableStyle source = ConstrainedImportedPreferenceStyle();
        source.MaxWidthFraction = source.PreferredWidthFraction = .5D;
        source.FrameWidthAdjustment = -1D;
        var nested = new TableBlock(new[] { new[] { new PdfTableCell(leftText), new PdfTableCell(rightText) } },
            PdfAlign.Left, source);
        var captured = new PdfTableCell(Array.Empty<PdfTextRun>(), null, blocks: new IPdfBlock[] { nested })
            .WithNamedDestination("ConstrainedImportedCell").WithNoWrap(false);
        source.PreserveAuthoredPreferencesAtWidthLimit = false;
        source.AutoFitColumnPreferredWidthPoints![0] = 105D;
        source.AutoFitCellPreferredWidthPoints![(0, 0)] = 105D;
        nested.Style!.PreserveAuthoredPreferencesAtWidthLimit = false;
        nested.Style.AutoFitColumnPreferredWidthPoints![0] = 105D;
        nested.Style.AutoFitCellPreferredWidthPoints![(0, 0)] = 105D;
        var outer = new PdfTableStyle {
            HeaderRowCount = 0, MaxWidth = global::OfficeIMO.OfficeLength.Points(216D), PreserveWidth = true,
            ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(216D) }, CellPaddingX = global::OfficeIMO.OfficeLength.Points(9D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D),
            BorderColor = null, FontSize = global::OfficeIMO.OfficeLength.Points(10D)
        };
        var document = PdfDocument.Create(ConstrainedImportedPreferenceOptions());
        document.Content.Table(new[] { new[] { captured } }, style: outer);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal(leftText + rightText, string.Concat(page.Letters.Select(letter => letter.Value)));
        double firstWidth = page.Letters[leftText.Length].StartBaseLine.X - page.Letters[0].StartBaseLine.X;
        double expected = 5D + (98.5D - 5D - 5D) * (45D - 5D) / ((45D - 5D) + (105D - 5D));
        Assert.Equal(expected, firstWidth, 3);
        Assert.Contains(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && Math.Abs(bounds.Value.Width - 98.5D) < .001D);
        double gridLeft = page.Letters[0].StartBaseLine.X - 2D;
        Assert.All(page.Letters, letter => Assert.True(letter.BoundingBox.Right <= gridLeft + 98.501D));
    }

    [Fact]
    public void ConstrainedImportedPreferencesKeepNormalAutomaticContentReflow() {
        double ordinary = MeasureConstrainedAutomaticRow(preserveAuthored: false);
        double optedIn = MeasureConstrainedAutomaticRow(preserveAuthored: true);
        Assert.Equal(ordinary, optedIn, 3);
        Assert.True(optedIn > 45D,
            "A later automatic cell whose word fits the complete frame must still grow beyond the earlier source preference.");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConstrainedImportedPreferencesRetainMandatoryPhysicalMinimumAtTheWidthLimit(bool inlineBox) {
        const string rightText = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string leftText;
        double minimumContentWidth;
        PdfTableCell leftCell;
        if (inlineBox) {
            leftText = "AB";
            minimumContentWidth = 60D;
            leftCell = new PdfTableCell(new[] {
                PdfTextRun.Normal("A"),
                PdfTextRun.Inline(new PdfInlineBox(minimumContentWidth, 8D, background: global::OfficeIMO.OfficeColor.FromNormalizedRgb(.2D, .5D, .8D))),
                PdfTextRun.Normal("B")
            });
        } else {
            leftText = "ABCDEFGHIJ";
            var metric = PdfDocument.Create(ConstrainedImportedPreferenceOptions());
            metric.Content.Text(leftText, style: new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(10D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) });
            using var metricPdf = PdfPigDocument.Open(metric.ToBytes());
            var measured = Assert.Single(metricPdf.GetPages()).Letters;
            minimumContentWidth = measured[measured.Count - 1].EndBaseLine.X - measured[0].StartBaseLine.X;
            leftCell = new PdfTableCell(leftText).WithNoWrap(true);
        }
        var style = ConstrainedImportedPreferenceStyle();
        var document = PdfDocument.Create(ConstrainedImportedPreferenceOptions());
        document.Content.Table(new[] { new[] { leftCell, new PdfTableCell(rightText) } }, style: style);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal(leftText + rightText, string.Concat(page.Letters.Select(letter => letter.Value)));
        double firstWidth = page.Letters[leftText.Length].StartBaseLine.X - page.Letters[0].StartBaseLine.X;
        Assert.True(firstWidth >= minimumContentWidth + 4D,
            "Compressing source preferences must retain each indivisible inline frame or NoWrap line plus authored padding.");
        double gridLeft = page.Letters[0].StartBaseLine.X - 2D;
        Assert.All(page.Letters, letter => Assert.True(letter.BoundingBox.Right <= gridLeft + 98.501D));
        if (inlineBox) {
            var box = Assert.Single(page.Paths.Where(path => path.IsFilled).Select(path => path.GetBoundingRectangle()),
                bounds => bounds.HasValue && Math.Abs(bounds.Value.Width - minimumContentWidth) < .001D &&
                    Math.Abs(bounds.Value.Height - 8D) < .001D);
            Assert.True(box!.Value.Left >= gridLeft + 2D - .001D);
            Assert.True(box.Value.Right <= gridLeft + firstWidth - 2D + .001D,
                "The actual painted inline frame must fit inside its own resolved cell content frame.");
        } else {
            Assert.All(page.Letters.Take(leftText.Length),
                letter => Assert.Equal(page.Letters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        }
    }

    private static double MeasureConstrainedAutomaticRow(bool preserveAuthored) {
        var style = ConstrainedImportedPreferenceStyle();
        style.PreferredWidth = style.MaxWidth = global::OfficeIMO.OfficeLength.Points(200D);
        style.PreserveAuthoredPreferencesAtWidthLimit = preserveAuthored;
        var document = PdfDocument.Create(ConstrainedImportedPreferenceOptions());
        // Row zero captures the source preferences. Row one has no absolute
        // cell entry and therefore retains the ordinary automatic-content rule.
        document.Content.Table(new[] {
            new[] { new PdfTableCell("R"), new PdfTableCell("S") },
            new[] { new PdfTableCell("ABCDEFGHIJ"), new PdfTableCell("B") }
        }, style: style);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal("RSABCDEFGHIJB", string.Concat(page.Letters.Select(letter => letter.Value)));
        var first = Assert.Single(page.Letters, letter => letter.Value == "R");
        var second = Assert.Single(page.Letters, letter => letter.Value == "S");
        return second.StartBaseLine.X - first.StartBaseLine.X;
    }

    private static PdfTableStyle ConstrainedImportedPreferenceStyle() => new() {
        HeaderRowCount = 0, AutoFitColumns = true, AutoFitWidthUsesContentMinimum = true,
        PreserveAuthoredPreferencesAtWidthLimit = true,
        PreserveWidth = true, AllowWidthBeyondFrame = true, ClipTextToCellBounds = true,
        PreferredWidth = global::OfficeIMO.OfficeLength.Points(150D), MaxWidth = global::OfficeIMO.OfficeLength.Points(98.5D),
        AutoFitColumnPreferredWidthPoints = new List<double?> { 45D, 105D },
        AutoFitCellPreferredWidthPoints = new Dictionary<(int Row, int Column), double> { [(0, 0)] = 45D, [(0, 1)] = 105D },
        ColumnMinWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(5D), global::OfficeIMO.OfficeLength.Points(5D) },
        CellPaddingX = global::OfficeIMO.OfficeLength.Points(2D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D), FontSize = global::OfficeIMO.OfficeLength.Points(10D), BorderColor = global::OfficeIMO.OfficeColor.Black, BorderWidth = global::OfficeIMO.OfficeLength.Points(.5D)
    };

    private static PdfCreateOptions ConstrainedImportedPreferenceOptions() => new() {
        PageWidth = global::OfficeIMO.OfficeLength.Points(340D), PageHeight = global::OfficeIMO.OfficeLength.Points(300D),
        MarginLeft = global::OfficeIMO.OfficeLength.Points(20D), MarginRight = global::OfficeIMO.OfficeLength.Points(20D), MarginTop = global::OfficeIMO.OfficeLength.Points(20D), MarginBottom = global::OfficeIMO.OfficeLength.Points(20D),
        DefaultFont = PdfStandardFont.Courier, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10D)
    };

}
