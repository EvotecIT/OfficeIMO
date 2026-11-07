using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfTableSourceFormattingTests {
    [Theory]
    [InlineData("TableGrid")]
    [InlineData("ListTable1LightAccent1")]
    public void SaveAsPdf_RecognizedTableStyleUsesSourcePaintInsteadOfPreset(string styleId) {
        using WordDocument document = CreateSourceTableDocument(styleId);
        SourceStyle(document, styleId).Append(new W.TableStyleProperties(
            new W.TableStyleConditionalFormattingTableCellProperties(
                new W.TableCellBorders(new W.BottomBorder {
                    Val = W.BorderValues.Single, Color = "FF0000", Size = 16U
                }),
                new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "FF0000" })) {
            Type = W.TableStyleOverrideValues.FirstRow
        });
        Assert.Empty(document.ValidateDocument());
        string before = SourceXml(document);

        using PdfPigDocument pdf = PdfPigDocument.Open(ConvertSourceTable(document));
        var page = Assert.Single(pdf.GetPages());
        var word = Assert.Single(page.GetWords(), item => item.Text == "SOURCE");
        Assert.All(word.Letters, letter => AssertRgb(letter.Color.ToRGBValues(), 0, 0, 0));
        Assert.DoesNotContain(page.Paths, path => path.IsFilled || path.IsStroked);
        Assert.Equal(before, SourceXml(document));
    }

    [Theory]
    [InlineData("table")]
    [InlineData("cell-style")]
    [InlineData("conditional")]
    [InlineData("direct")]
    public void SaveAsPdf_TablePaintUsesDocumentThemeBeforeCachedRgb(string layer) {
        using WordDocument document = CreateSourceTableDocument("TableGrid", firstRow: layer == "conditional");
        W.Style style = SourceStyle(document, "TableGrid");
        W.TableCell cell = document.OpenXmlDocument.MainDocumentPart!.Document.Descendants<W.TableCell>().Single();
        if (layer is "table" or "cell-style") {
            style.StyleRunProperties = new W.StyleRunProperties(ThemedTextColor());
            style.StyleTableProperties = new W.StyleTableProperties(ThemedTableBorders(), layer == "table"
                ? ThemedFill() : new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "00FFFF" });
            if (layer == "cell-style") style.StyleTableCellProperties = new W.StyleTableCellProperties(ThemedFill());
        } else if (layer == "conditional") {
            style.Append(new W.TableStyleProperties(
                new W.RunPropertiesBaseStyle(ThemedTextColor()),
                new W.TableStyleConditionalFormattingTableCellProperties(ThemedCellBorders(), ThemedFill())) {
                Type = W.TableStyleOverrideValues.FirstRow
            });
        } else {
            cell.TableCellProperties!.TableCellBorders = ThemedCellBorders();
            cell.TableCellProperties.Shading = ThemedFill();
            cell.Descendants<W.Run>().Single().RunProperties!.Color = ThemedTextColor();
        }
        Assert.Empty(document.ValidateDocument());
        string before = SourceXml(document);

        using PdfPigDocument pdf = PdfPigDocument.Open(ConvertSourceTable(document));
        var page = Assert.Single(pdf.GetPages());
        var word = Assert.Single(page.GetWords(), item => item.Text == "SOURCE");
        Assert.All(word.Letters, letter => AssertRgb(letter.Color.ToRGBValues(), 217, 87, 22));
        var fills = page.Paths.Where(path => path.IsFilled && MatchesRgb(path.FillColor.ToRGBValues(), 18, 164, 94)).ToArray();
        Assert.NotEmpty(fills);
        // The source cell is 3600 by 800 twips: 180 by 40 points.
        Assert.All(fills, fill => {
            var bounds = fill.GetBoundingRectangle()!.Value;
            Assert.InRange(bounds.Width, 178D, 182D);
            Assert.InRange(bounds.Height, 38D, 42D);
            Assert.True(bounds.Left <= word.BoundingBox.Left && bounds.Right >= word.BoundingBox.Right &&
                bounds.Bottom <= word.BoundingBox.Bottom && bounds.Top >= word.BoundingBox.Top);
        });
        var fillBounds = fills[0].GetBoundingRectangle()!.Value;
        var borders = page.Paths.Where(path => path.IsStroked).ToArray();
        Assert.NotEmpty(borders);
        Assert.All(borders, border => {
            AssertRgb(border.StrokeColor.ToRGBValues(), 138, 52, 196);
            Assert.Equal(2D, border.LineWidth, 3);
            var bounds = border.GetBoundingRectangle()!.Value;
            Assert.InRange(bounds.Left, fillBounds.Left - 1D, fillBounds.Right + 1D);
            Assert.InRange(bounds.Right, fillBounds.Left - 1D, fillBounds.Right + 1D);
            Assert.InRange(bounds.Bottom, fillBounds.Bottom - 1D, fillBounds.Top + 1D);
            Assert.InRange(bounds.Top, fillBounds.Bottom - 1D, fillBounds.Top + 1D);
        });
        Assert.Equal(before, SourceXml(document));
    }

    [Fact]
    public void SaveAsPdf_LiteralCellBorderEditReplacesThemePaint() {
        using WordDocument document = CreateSourceTableDocument("TableGrid");
        W.TableCell cellXml = document.OpenXmlDocument.MainDocumentPart!.Document.Descendants<W.TableCell>().Single();
        cellXml.TableCellProperties!.TableCellBorders = new W.TableCellBorders(new W.TopBorder {
            Val = W.BorderValues.Single, Color = "0000FF", Size = 16U,
            ThemeColor = W.ThemeColorValues.Accent2, ThemeTint = "80", ThemeShade = "80"
        });
        document.Tables.Single().Rows[0].Cells[0].Borders.TopColor = OfficeColor.Red;
        W.TopBorder editedBorder = cellXml.TableCellProperties.TableCellBorders.TopBorder!;
        Assert.Null(editedBorder.ThemeColor);
        Assert.Null(editedBorder.ThemeTint);
        Assert.Null(editedBorder.ThemeShade);
        Assert.Empty(document.ValidateDocument());
        string before = SourceXml(document);

        using PdfPigDocument pdf = PdfPigDocument.Open(ConvertSourceTable(document));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains(page.GetWords(), word => word.Text == "SOURCE");
        var border = Assert.Single(page.Paths, path => path.IsStroked);
        AssertRgb(border.StrokeColor.ToRGBValues(), 255, 0, 0);
        Assert.Equal(2D, border.LineWidth, 3);
        var bounds = border.GetBoundingRectangle()!.Value;
        Assert.InRange(bounds.Width, 178D, 182D);
        Assert.InRange(bounds.Height, 0D, .001D);
        Assert.Equal(before, SourceXml(document));
    }

    [Theory]
    [InlineData("first", false)]
    [InlineData("first", true)]
    [InlineData("last", false)]
    [InlineData("configured", false)]
    public void SaveAsPdf_ExplicitNoFillClearsInheritedRowPaint(string role, bool clear) {
        bool configuredHeader = role == "configured";
        bool lastRow = role == "last";
        string styleId = configuredHeader ? "TableNormal" : "TableGrid";
        using WordDocument document = CreateSourceTableDocument(styleId, firstRow: !lastRow);
        if (!configuredHeader) {
            SourceStyle(document, styleId).Append(new W.TableStyleProperties(
                new W.TableStyleConditionalFormattingTableCellProperties(ThemedFill())) {
                Type = lastRow ? W.TableStyleOverrideValues.LastRow : W.TableStyleOverrideValues.FirstRow
            });
        }
        W.Table table = document.OpenXmlDocument.MainDocumentPart!.Document.Descendants<W.Table>().Single();
        table.GetFirstChild<W.TableProperties>()!.TableLook!.LastRow = lastRow;
        table.GetFirstChild<W.TableGrid>()!.GetFirstChild<W.GridColumn>()!.Width = "1800";
        table.GetFirstChild<W.TableGrid>()!.Append(new W.GridColumn { Width = "1800" });
        W.TableRow row = table.GetFirstChild<W.TableRow>()!;
        W.TableCell first = row.GetFirstChild<W.TableCell>()!;
        first.TableCellProperties!.TableCellWidth!.Width = "1800";
        first.Descendants<W.Text>().Single().Text = "CLEAR";
        W.TableCell inherited = (W.TableCell)first.CloneNode(true);
        inherited.Descendants<W.Text>().Single().Text = "INHERITED";
        row.Append(inherited);
        first.TableCellProperties.Shading = new W.Shading {
            Val = clear ? W.ShadingPatternValues.Clear : W.ShadingPatternValues.Nil,
            Fill = clear ? "auto" : "FF0000"
        };
        Assert.Empty(document.ValidateDocument());
        string before = SourceXml(document);
        PdfTableStyle? configuredStyle = configuredHeader ? new PdfTableStyle {
            HeaderFill = OfficeColor.FromRgb(18, 164, 94), HeaderBold = false,
            FooterFill = null, RowStripeFill = null, BorderColor = null, BorderWidth = OfficeLength.Zero
        } : null;

        using PdfPigDocument pdf = PdfPigDocument.Open(ConvertSourceTable(document, configuredStyle));
        var page = Assert.Single(pdf.GetPages());
        var inheritedWord = Assert.Single(page.GetWords(), word => word.Text == "INHERITED");
        var clearWord = Assert.Single(page.GetWords(), word => word.Text == "CLEAR");
        double inheritedX = (inheritedWord.BoundingBox.Left + inheritedWord.BoundingBox.Right) / 2D;
        double inheritedY = (inheritedWord.BoundingBox.Bottom + inheritedWord.BoundingBox.Top) / 2D;
        double clearX = (clearWord.BoundingBox.Left + clearWord.BoundingBox.Right) / 2D;
        double clearY = (clearWord.BoundingBox.Bottom + clearWord.BoundingBox.Top) / 2D;
        Assert.Contains(page.Paths, path => path.IsFilled && MatchesRgb(path.FillColor.ToRGBValues(), 18, 164, 94) &&
            path.GetBoundingRectangle() is { } bounds && bounds.Left < inheritedX && bounds.Right > inheritedX &&
            bounds.Bottom < inheritedY && bounds.Top > inheritedY);
        Assert.DoesNotContain(page.Paths, path => path.IsFilled && path.GetBoundingRectangle() is { } bounds &&
            bounds.Left < clearX && bounds.Right > clearX && bounds.Bottom < clearY && bounds.Top > clearY);
        Assert.Equal(before, SourceXml(document));
    }

    private static WordDocument CreateSourceTableDocument(string styleId, bool firstRow = false) {
        WordDocument document = WordDocument.Create();
        document.PageSetup.Size.Width = OfficeLength.Points(240D);
        document.PageSetup.Size.Height = OfficeLength.Points(300D);
        document.PageSetup.Margins.Left = document.PageSetup.Margins.Right = OfficeLength.Points(30D);
        document.PageSetup.Margins.Top = document.PageSetup.Margins.Bottom = OfficeLength.Points(30D);
        var main = document.OpenXmlDocument.MainDocumentPart!;
        W.Styles styles = main.StyleDefinitionsPart!.Styles!;
        foreach (W.Style existing in styles.Elements<W.Style>().Where(item => item.StyleId?.Value == styleId).ToArray())
            existing.Remove();
        styles.Append(new W.Style(new W.StyleName { Val = styleId }) { Type = W.StyleValues.Table, StyleId = styleId });
        styles.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.Color { Val = "000000" }, new W.FontSize { Val = "24" })));
        styles.Elements<W.Style>().Single(item => item.StyleId?.Value == "Normal").StyleRunProperties = new W.StyleRunProperties();
        A.ColorScheme scheme = main.ThemePart!.Theme.ThemeElements!.ColorScheme!;
        scheme.Accent1Color = new A.Accent1Color(new A.RgbColorModelHex { Val = "12A45E" });
        scheme.Accent2Color = new A.Accent2Color(new A.RgbColorModelHex { Val = "8A34C4" });
        scheme.Accent3Color = new A.Accent3Color(new A.RgbColorModelHex { Val = "D95716" });
        W.Body body = main.Document.Body!;
        body.RemoveAllChildren<W.Paragraph>();
        var table = new W.Table(
            new W.TableProperties(
                new W.TableStyle { Val = styleId },
                new W.TableWidth { Type = W.TableWidthUnitValues.Dxa, Width = "3600" },
                new W.TableLayout { Type = W.TableLayoutValues.Fixed },
                new W.TableLook {
                    FirstRow = firstRow, LastRow = false, FirstColumn = false, LastColumn = false,
                    NoHorizontalBand = true, NoVerticalBand = true
                }),
            new W.TableGrid(new W.GridColumn { Width = "3600" }),
            new W.TableRow(
                new W.TableRowProperties(new W.TableRowHeight { Val = 800U, HeightType = W.HeightRuleValues.Exact }),
                new W.TableCell(
                    new W.TableCellProperties(new W.TableCellWidth { Type = W.TableWidthUnitValues.Dxa, Width = "3600" }),
                    new W.Paragraph(
                        new W.ParagraphProperties(new W.SpacingBetweenLines { Before = "0", After = "0" }),
                        new W.Run(new W.RunProperties(new W.FontSize { Val = "24" }), new W.Text("SOURCE"))))));
        W.SectionProperties? section = body.GetFirstChild<W.SectionProperties>();
        if (section == null) body.Append(table);
        else body.InsertBefore(table, section);
        return document;
    }

    private static W.Style SourceStyle(WordDocument document, string styleId) =>
        document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<W.Style>()
            .Single(style => style.StyleId?.Value == styleId);

    private static W.Shading ThemedFill() => new() {
        Val = W.ShadingPatternValues.Clear, Fill = "FFFF00", ThemeFill = W.ThemeColorValues.Accent1
    };

    private static W.Color ThemedTextColor() => new() { Val = "0000FF", ThemeColor = W.ThemeColorValues.Accent3 };

    private static W.TableBorders ThemedTableBorders() => new(
        ThemedBorder<W.TopBorder>(), ThemedBorder<W.LeftBorder>(), ThemedBorder<W.BottomBorder>(), ThemedBorder<W.RightBorder>());

    private static W.TableCellBorders ThemedCellBorders() => new(
        ThemedBorder<W.TopBorder>(), ThemedBorder<W.LeftBorder>(), ThemedBorder<W.BottomBorder>(), ThemedBorder<W.RightBorder>());

    private static T ThemedBorder<T>() where T : W.BorderType, new() => new() {
        Val = W.BorderValues.Single, Color = "FF0000", ThemeColor = W.ThemeColorValues.Accent2, Size = 16U
    };

    private static byte[] ConvertSourceTable(WordDocument document, PdfTableStyle? configuredStyle = null) {
        var options = new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        if (configuredStyle != null) options.PdfOptions = new PdfCreateOptions { DefaultTableStyle = configuredStyle };
        using var model = document.ToPdf(options);
        byte[] bytes = model.ToBytes(new PdfSaveOptions { AllowLoss = true });
        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        return bytes;
    }

    private static void AssertRgb((double r, double g, double b) actual, byte red, byte green, byte blue) {
        Assert.Equal(((int)red, (int)green, (int)blue),
            ((int)Math.Round(actual.r * 255D), (int)Math.Round(actual.g * 255D), (int)Math.Round(actual.b * 255D)));
    }

    private static bool MatchesRgb((double r, double g, double b) actual, byte red, byte green, byte blue) =>
        (int)Math.Round(actual.r * 255D) == red && (int)Math.Round(actual.g * 255D) == green &&
        (int)Math.Round(actual.b * 255D) == blue;

    private static string SourceXml(WordDocument document) {
        var main = document.OpenXmlDocument.MainDocumentPart!;
        return main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml + main.ThemePart!.Theme.OuterXml +
            main.DocumentSettingsPart!.Settings!.OuterXml;
    }
}
