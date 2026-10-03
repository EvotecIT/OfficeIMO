using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, "Arial")]
    [InlineData(false, false, "Times New Roman")]
    [InlineData(false, true, "Arial")]
    [InlineData(false, true, "Times New Roman")]
    [InlineData(true, false, "Arial")]
    [InlineData(true, false, "Times New Roman")]
    [InlineData(true, true, "Arial")]
    [InlineData(true, true, "Times New Roman")]
    public void SaveAsPdf_InheritedAutoSpacingUsesTheEffectiveParagraphFont(bool nativeDoc, bool styleSpacing, string family) {
        double[] direct = InheritedLeadingBaselines(nativeDoc, styleSpacing, family, directSpacing: true);
        double[] inherited = InheritedLeadingBaselines(nativeDoc, styleSpacing, family, directSpacing: false);
        Assert.Equal(3, direct.Length);
        Assert.Equal(3, inherited.Length);
        Assert.True(direct[0] - direct[1] > 12);
        for (int line = 1; line < 3; line++) {
            Assert.Equal(direct[line - 1] - direct[line], inherited[line - 1] - inherited[line], precision: 3);
        }
    }

    [Theory]
    [InlineData(false, false, WordLineSpacingRule.Exact)]
    [InlineData(false, true, WordLineSpacingRule.Exact)]
    [InlineData(true, false, WordLineSpacingRule.Exact)]
    [InlineData(true, true, WordLineSpacingRule.Exact)]
    [InlineData(false, false, WordLineSpacingRule.AtLeast)]
    [InlineData(false, true, WordLineSpacingRule.AtLeast)]
    [InlineData(true, false, WordLineSpacingRule.AtLeast)]
    [InlineData(true, true, WordLineSpacingRule.AtLeast)]
    public void SaveAsPdf_InheritedFixedSpacingRetainsItsPointHeight(bool nativeDoc, bool styleSpacing, WordLineSpacingRule rule) {
        double[] baselines = InheritedLeadingBaselines(nativeDoc, styleSpacing, "Arial", false, lineSpacing: 440, rule: rule);
        Assert.Equal(3, baselines.Length);
        Assert.Equal(22D, baselines[0] - baselines[1], precision: 3);
        Assert.Equal(22D, baselines[1] - baselines[2], precision: 3);
    }

    [Theory]
    [InlineData("columns")]
    [InlineData("list")]
    [InlineData("heading")]
    [InlineData("derived")]
    public void SaveAsPdf_InheritedAutoSpacingSurvivesOtherParagraphRoutes(string route) {
        double[] direct = InheritedLeadingBaselines(false, true, "Arial", true, route: route);
        double[] inherited = InheritedLeadingBaselines(false, true, "Arial", false, route: route);
        Assert.Equal(3, direct.Length);
        Assert.Equal(3, inherited.Length);
        for (int line = 1; line < 3; line++) {
            Assert.Equal(18D, inherited[line - 1] - inherited[line], precision: 3);
            Assert.Equal(direct[line - 1] - direct[line], inherited[line - 1] - inherited[line], precision: 3);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_AuthoredLineWithoutRuleMeansAutoEvenOverExactDefaults(bool nativeDoc) {
        double[] baselines = InheritedLeadingBaselines(nativeDoc, false, "Arial", false, omitRule: true);
        Assert.Equal(18D, baselines[0] - baselines[1], precision: 3);
        Assert.Equal(18D, baselines[1] - baselines[2], precision: 3);
    }

    [Theory]
    [InlineData("document")]
    [InlineData("paragraph")]
    [InlineData("table")]
    [InlineData("conditional")]
    public void SaveAsPdf_TableInheritedAutoSpacingUsesItsFinalFont(string tier) {
        double[] direct = InheritedLeadingBaselines(false, tier == "paragraph", "Arial", true, route: tier);
        double[] inherited = InheritedLeadingBaselines(false, tier == "paragraph", "Arial", false, route: tier);
        Assert.Equal(3, inherited.Length);
        Assert.Equal(3, direct.Length);
        for (int line = 1; line < 3; line++) {
            Assert.Equal(18D, inherited[line - 1] - inherited[line], precision: 3);
            Assert.Equal(direct[line - 1] - direct[line], inherited[line - 1] - inherited[line], precision: 3);
        }
    }

    private static double[] InheritedLeadingBaselines(bool nativeDoc, bool styleSpacing, string family, bool directSpacing,
        int lineSpacing = 360, WordLineSpacingRule rule = WordLineSpacingRule.Auto, string route = "body", bool omitRule = false) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(
            new RunPropertiesDefault(new RunPropertiesBaseStyle(new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }, new FontSize { Val = "24" })),
            new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(new SpacingBetweenLines {
                Line = styleSpacing || route is "table" or "conditional" ? "240" : lineSpacing.ToString(System.Globalization.CultureInfo.InvariantCulture),
                LineRule = omitRule ? LineSpacingRuleValues.Exact : rule.ToOpenXml(), After = "0"
            })));
        styles.Append(new Style(new StyleName { Val = "Inherited leading" },
            new StyleRunProperties(new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }, new FontSize { Val = "24" }),
            new StyleParagraphProperties(styleSpacing ? new SpacingBetweenLines {
                Line = lineSpacing.ToString(System.Globalization.CultureInfo.InvariantCulture), LineRule = rule.ToOpenXml()
            } : new SpacingBetweenLines { After = "0" })) {
            Type = StyleValues.Paragraph, StyleId = "InheritedLeading", CustomStyle = true
        });
        WordParagraph paragraph;
        if (route is "document" or "paragraph" or "table" or "conditional") {
            var table = source.AddTable(1, 1, WordTableStyle.TableGrid);
            if (route == "document") table._tableProperties!.TableStyle = null;
            paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "A\nB\nC";
            if (route is "table" or "conditional") {
                var properties = new StyleParagraphProperties(new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto });
                var tableStyle = new Style(new StyleName { Val = "Inherited table leading" },
                    new StyleRunProperties(new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" })) {
                    Type = StyleValues.Table, StyleId = "InheritedTableLeading", CustomStyle = true
                };
                if (route == "table") tableStyle.Append(properties);
                else tableStyle.Append(new TableStyleProperties(properties) { Type = TableStyleOverrideValues.FirstRow });
                styles.Append(tableStyle);
                table._tableProperties!.TableStyle = new TableStyle { Val = "InheritedTableLeading" };
                table.ConditionalFormattingFirstRow = true;
            }
        } else if (route == "list") {
            var list = source.AddCustomList();
            list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
            paragraph = list.AddItem("A\nB\nC");
        } else {
            paragraph = source.AddParagraph("A\nB\nC");
            if (route == "columns") source.Sections[0].ColumnCount = 2;
            if (route == "heading") styles.Elements<Style>().Single(style => style.StyleId == "InheritedLeading")
                .StyleParagraphProperties!.Append(new OutlineLevel { Val = 0 });
        }
        if (route == "derived") {
            styles.Append(new Style(new StyleName { Val = "Derived leading" }, new BasedOn { Val = "InheritedLeading" },
                new StyleRunProperties(new RunFonts { Ascii = family, HighAnsi = family }),
                new StyleParagraphProperties(new SpacingBetweenLines { After = "0" })) {
                Type = StyleValues.Paragraph, StyleId = "DerivedLeading", CustomStyle = true
            });
            paragraph.SetStyleId("DerivedLeading");
        } else {
            paragraph.SetStyleId("InheritedLeading");
            paragraph.FontFamily = family;
        }
        paragraph.FontSize = 12;
        paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        if (directSpacing || omitRule) { paragraph.LineSpacing = lineSpacing; paragraph.LineSpacingRule = omitRule ? null : rule; }
        string originalStyles = styles.OuterXml;
        byte[] documentBytes = nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes();
        Assert.Equal(originalStyles, styles.OuterXml);
        using WordDocument document = WordDocument.Load(new System.IO.MemoryStream(documentBytes));
        var pdfOptions = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        pdfOptions.EmbedStandardFont(PdfStandardFont.Helvetica, ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray()), "OfficeIMO-Portable-Regular");
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = pdfOptions });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "B" or "C").ToArray();
        Assert.Equal("ABC", string.Concat(letters.Select(letter => letter.Value)));
        return letters.Select(letter => letter.StartBaseLine.Y).ToArray();
    }
}
