using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, (short)40)]
    [InlineData(false, (short)0)]
    [InlineData(true, (short)40)]
    public void SaveAsPdf_NestedTableUsesItsPhysicalCellFrameAcrossCompatibilityModes(bool percentage, short childMarginTwips) {
        var modern = MeasureNestedCompatibilityFrame(WordCompatibilityMode.Word2013, percentage, childMarginTwips);
        var legacy = MeasureNestedCompatibilityFrame(WordCompatibilityMode.Word2010, percentage, childMarginTwips);

        Assert.True(legacy.AnchorX < modern.AnchorX,
            "The outer body table must retain its own legacy text-margin placement.");
        Assert.Equal(modern.RelativeChildX, legacy.RelativeChildX, precision: 2);
        Assert.Equal(modern.ColumnDistance, legacy.ColumnDistance, precision: 2);
    }

    [Fact]
    public void SaveAsPdf_NestedPercentageFrameIncludesItsOwnSeparateBorderSpacing() {
        var percentage = MeasureNestedSeparateDeclaredFrame(percentage: true);
        var absolute = MeasureNestedSeparateDeclaredFrame(percentage: false);
        Assert.Equal(absolute.RelativeChildX, percentage.RelativeChildX, precision: 2);
        Assert.Equal(absolute.ColumnDistance, percentage.ColumnDistance, precision: 2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_CompressedNestedFixedGridRetainsPhysicalCellMargins(bool separateBorders) {
        const short additionalMarginTwips = 40;
        const short initialMarginTwips = 40;
        // Start above the mandatory half-border inset so added source margins
        // are physical growth rather than a change in border-floor admission.
        double originalColumnDistance = MeasureCompressedNestedFixedColumnDistance(initialMarginTwips, separateBorders);
        double enlargedColumnDistance = MeasureCompressedNestedFixedColumnDistance((short)(initialMarginTwips + additionalMarginTwips), separateBorders);
        // Both authored column boxes grow by their two added margins, and the
        // declared table frame grows by exactly the same total. The text-space
        // compression is therefore unchanged; margins remain physical distances.
        double addedMarginsInFirstCell = 2D * additionalMarginTwips / 20D;
        Assert.Equal(addedMarginsInFirstCell, enlargedColumnDistance - originalColumnDistance, precision: 2);
    }

    private static double MeasureCompressedNestedFixedColumnDistance(short childMarginTwips, bool separateBorders) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(outer, 6000, 6000);
        WordTableCell parent = outer.Rows[0].Cells[0];
        parent.MarginLeftWidth = parent.MarginRightWidth = 180;
        ConfigureNestedLayoutParagraph(parent.Paragraphs[0], "FRAMEANCHOR");
        WordTable inner = parent.AddTable(1, 2, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(inner, 2000 + 4 * childMarginTwips,
            900 + 2 * childMarginTwips, 2100 + 2 * childMarginTwips);
        inner._tableProperties!.TableIndentation = new W.TableIndentation { Width = 0, Type = W.TableWidthUnitValues.Dxa };
        if (separateBorders)
            inner._tableProperties.TableCellSpacing = new W.TableCellSpacing { Width = "60", Type = W.TableWidthUnitValues.Dxa };
        foreach (WordTableCell cell in inner.Rows[0].Cells) cell.MarginLeftWidth = cell.MarginRightWidth = childMarginTwips;
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[0].Paragraphs[0], "X");
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[1].Paragraphs[0], "Y");
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string original = main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml + main.DocumentSettingsPart!.Settings!.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new PageSize(400D, 500D), Margins = PageMargins.Uniform(30D),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        Assert.Single(words, word => word.Text == "FRAMEANCHOR");
        var left = Assert.Single(words, word => word.Text == "X");
        var right = Assert.Single(words, word => word.Text == "Y");
        Assert.Equal(left.BoundingBox.Bottom, right.BoundingBox.Bottom, precision: 2);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeNestedTableLayoutApproximated");
        Assert.Equal(original, main.Document.OuterXml + main.StyleDefinitionsPart.Styles.OuterXml + main.DocumentSettingsPart.Settings.OuterXml);
        return right.BoundingBox.Left - left.BoundingBox.Left;
    }

    [Theory]
    [InlineData(WordCompatibilityMode.Word2010)]
    [InlineData(WordCompatibilityMode.Word2013)]
    public void SaveAsPdf_NestedPercentageAutoFitRetainsConstrainedAuthoredCellPreferences(WordCompatibilityMode mode) {
        var fixedGrid = MeasureNestedCappedAutoFitGrid(mode, automatic: false);
        var automaticGrid = MeasureNestedCappedAutoFitGrid(mode, automatic: true);
        Assert.Equal(fixedGrid.ColumnDistance, automaticGrid.ColumnDistance, precision: 2);
        Assert.Equal(fixedGrid.Glyphs, automaticGrid.Glyphs);
    }

    private static (double ColumnDistance, (string Text, double X, double Y)[] Glyphs) MeasureNestedCappedAutoFitGrid(
        WordCompatibilityMode mode, bool automatic) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(outer, 4320, 4320);
        WordTableCell parent = outer.Rows[0].Cells[0];
        parent.MarginLeftWidth = parent.MarginRightWidth = 180;
        ConfigureNestedLayoutParagraph(parent.Paragraphs[0], "FRAMEANCHOR");
        WordTable inner = parent.AddTable(2, 2, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(inner, 3000, 900, 2100);
        inner.WidthType = WordTableWidthUnit.Pct;
        inner.Width = 2500;
        inner.LayoutMode = automatic ? WordTableLayoutMode.AutoFit : WordTableLayoutMode.Fixed;
        inner._tableProperties!.TableIndentation = new W.TableIndentation { Width = 0, Type = W.TableWidthUnitValues.Dxa };
        inner._tableProperties.TableBorders = new W.TableBorders(
            new W.TopBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.LeftBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.BottomBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.RightBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 8U });
        foreach (WordTableRow row in inner.Rows)
            foreach (WordTableCell cell in row.Cells) cell.MarginLeftWidth = cell.MarginRightWidth = 40;
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[0].Paragraphs[0], "X");
        ConfigureNestedLayoutParagraph(inner.Rows[1].Cells[0].Paragraphs[0], "Y");
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[1].Paragraphs[0], "R1 A characteristically detailed description.");
        ConfigureNestedLayoutParagraph(inner.Rows[1].Cells[1].Paragraphs[0], "R2 A characteristically detailed description.");
        ConfigureNestedLayoutParagraph(document.AddParagraph(), "BODYAFTER");
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string original = main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml + main.DocumentSettingsPart!.Settings!.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new PageSize(400D, 500D), Margins = PageMargins.Uniform(30D),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var words = page.GetWords().ToArray();
        var left = Assert.Single(words, word => word.Text == "X");
        var right = Assert.Single(words, word => word.Text == "R1");
        var letters = page.Letters.ToArray();
        Assert.Equal("FRAMEANCHORXR1Acharacteristicallydetaileddescription.YR2Acharacteristicallydetaileddescription.BODYAFTER",
            string.Concat(letters.Select(letter => letter.Value)).Replace(" ", ""));
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeNestedTableLayoutApproximated");
        Assert.Equal(original, main.Document.OuterXml + main.StyleDefinitionsPart.Styles.OuterXml + main.DocumentSettingsPart.Settings.OuterXml);
        return (right.BoundingBox.Left - left.BoundingBox.Left, letters.Select(letter =>
            (letter.Value, Math.Round(letter.BoundingBox.Left, 2), Math.Round(letter.BoundingBox.Bottom, 2))).ToArray());
    }

    private static (double RelativeChildX, double ColumnDistance) MeasureNestedSeparateDeclaredFrame(bool percentage) {
        const int parentWidthTwips = 6000;
        const short parentMarginTwips = 180;
        const uint borderSizeEighthsOfPoint = 8U;
        const int borderWidthTwips = 20;
        const int percentageFiftieths = 2500;
        // This absolute counterpart is derived from the same authored cell frame
        // and border, independently of the table's own separate-spacing envelope.
        int declaredWidthTwips = (parentWidthTwips - 2 * parentMarginTwips - borderWidthTwips) * percentageFiftieths / 5000;
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(outer, parentWidthTwips, parentWidthTwips);
        WordTableCell parent = outer.Rows[0].Cells[0];
        parent.MarginLeftWidth = parent.MarginRightWidth = parentMarginTwips;
        ConfigureNestedLayoutParagraph(parent.Paragraphs[0], "FRAMEANCHOR");
        WordTable inner = parent.AddTable(1, 2, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(inner, declaredWidthTwips, 900, 2100);
        inner._tableProperties!.TableIndentation = new W.TableIndentation { Width = 0, Type = W.TableWidthUnitValues.Dxa };
        inner._tableProperties.TableCellSpacing = new W.TableCellSpacing { Width = "60", Type = W.TableWidthUnitValues.Dxa };
        inner._tableProperties.TableBorders = new W.TableBorders(
            new W.TopBorder { Val = W.BorderValues.Single, Size = borderSizeEighthsOfPoint },
            new W.LeftBorder { Val = W.BorderValues.Single, Size = borderSizeEighthsOfPoint },
            new W.BottomBorder { Val = W.BorderValues.Single, Size = borderSizeEighthsOfPoint },
            new W.RightBorder { Val = W.BorderValues.Single, Size = borderSizeEighthsOfPoint },
            new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = borderSizeEighthsOfPoint },
            new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = borderSizeEighthsOfPoint });
        if (percentage) { inner.WidthType = WordTableWidthUnit.Pct; inner.Width = percentageFiftieths; }
        foreach (WordTableCell cell in inner.Rows[0].Cells) cell.MarginLeftWidth = cell.MarginRightWidth = 40;
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[0].Paragraphs[0], "LEFT");
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[1].Paragraphs[0], "RIGHT");
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string original = main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml + main.DocumentSettingsPart!.Settings!.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new PageSize(400D, 500D), Margins = PageMargins.Uniform(30D),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var anchor = Assert.Single(words, word => word.Text == "FRAMEANCHOR");
        var left = Assert.Single(words, word => word.Text == "LEFT");
        var right = Assert.Single(words, word => word.Text == "RIGHT");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeNestedTableLayoutApproximated");
        Assert.Equal(original, main.Document.OuterXml + main.StyleDefinitionsPart.Styles.OuterXml + main.DocumentSettingsPart.Settings.OuterXml);
        return (left.BoundingBox.Left - anchor.BoundingBox.Left, right.BoundingBox.Left - left.BoundingBox.Left);
    }

    private static (double AnchorX, double RelativeChildX, double ColumnDistance) MeasureNestedCompatibilityFrame(
        WordCompatibilityMode mode, bool percentage, short childMarginTwips) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(outer, 6000, 6000);
        WordTableCell parent = outer.Rows[0].Cells[0];
        parent.MarginLeftWidth = parent.MarginRightWidth = 180;
        ConfigureNestedLayoutParagraph(parent.Paragraphs[0], "FRAMEANCHOR");
        WordTable inner = parent.AddTable(1, 2, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(inner, 3000, 900, 2100);
        inner._tableProperties!.TableIndentation = new W.TableIndentation { Width = 0, Type = W.TableWidthUnitValues.Dxa };
        if (percentage) {
            inner.WidthType = WordTableWidthUnit.Pct;
            inner.Width = 2500;
        }
        foreach (WordTableCell cell in inner.Rows[0].Cells)
            cell.MarginLeftWidth = cell.MarginRightWidth = childMarginTwips;
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[0].Paragraphs[0], "LEFT");
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[1].Paragraphs[0], "RIGHT");
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string SourceXml() => main.Document.OuterXml + main.StyleDefinitionsPart?.Styles?.OuterXml +
            main.DocumentSettingsPart?.Settings?.OuterXml;
        string original = SourceXml();
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new PageSize(400D, 500D), Margins = PageMargins.Uniform(30D),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var anchor = Assert.Single(words, word => word.Text == "FRAMEANCHOR");
        var left = Assert.Single(words, word => word.Text == "LEFT");
        var right = Assert.Single(words, word => word.Text == "RIGHT");
        Assert.Equal(left.BoundingBox.Bottom, right.BoundingBox.Bottom, precision: 2);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeNestedTableLayoutApproximated");
        Assert.Equal(original, SourceXml());
        return (anchor.BoundingBox.Left, left.BoundingBox.Left - anchor.BoundingBox.Left,
            right.BoundingBox.Left - left.BoundingBox.Left);
    }
}
