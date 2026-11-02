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
    [InlineData("top")]
    [InlineData("center")]
    [InlineData("bottom")]
    [InlineData("baseline")]
    public void SaveAsPdf_ExactLineAlignmentKeepsUniformListMarkersWithTheirBody(string alignment) {
        const string family = "OfficeIMO list device alignment";
        using WordDocument document = WordDocument.Create();
        var list = document.AddCustomList();
        var level = new WordListLevel(WordListLevelKind.DecimalDot) { LevelText = "A", LevelSuffix = WordListLevelSuffix.Tab };
        level.OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
            new RunFonts { Ascii = family, HighAnsi = family }, new FontSize { Val = "16" });
        list.Numbering.AddLevel(level);
        var paragraph = list.AddItem("B");
        paragraph.FontFamily = family; paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(8);
        paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(30));
        paragraph._paragraph.ParagraphProperties!.TextAlignment = new TextAlignment { Val = new VerticalTextAlignmentValues(alignment) };
        var options = new PdfCreateOptions();
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics()));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(Assert.Single(letters, x => x.Value == "A").StartBaseLine.Y,
            Assert.Single(letters, x => x.Value == "B").StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false, false, "top", 7.08D)]
    [InlineData(false, true, "top", 7.08D)]
    [InlineData(true, false, "top", 7.08D)]
    [InlineData(true, true, "top", 7.08D)]
    [InlineData(false, false, "center", 2.64D)]
    [InlineData(false, true, "center", 2.64D)]
    [InlineData(true, false, "center", 2.64D)]
    [InlineData(true, true, "center", 2.64D)]
    [InlineData(false, false, "bottom", -.96D)]
    [InlineData(false, true, "bottom", -.96D)]
    [InlineData(true, false, "bottom", -.96D)]
    [InlineData(true, true, "bottom", -.96D)]
    [InlineData(false, false, "baseline", 0D)]
    [InlineData(false, true, "baseline", 0D)]
    [InlineData(true, false, "baseline", 0D)]
    [InlineData(true, true, "baseline", 0D)]
    [InlineData(false, false, "auto", 0D)]
    [InlineData(false, true, "auto", 0D)]
    [InlineData(true, false, "auto", 0D)]
    [InlineData(true, true, "auto", 0D)]
    public void SaveAsPdf_ExactLineAlignmentRetainsPerRunMetrics(bool nativeDoc, bool table, string alignment, double difference) {
        using var pdf = PdfPigDocument.Open(CreateLineAlignmentPdf("direct", alignment, nativeDoc, table));
        AssertLineAlignmentDifference(pdf, difference);
    }

    [Theory]
    [InlineData("document", false, 7.08D)]
    [InlineData("document", true, 7.08D)]
    [InlineData("style", false, 7.08D)]
    [InlineData("style", true, 7.08D)]
    [InlineData("table", false, 7.08D)]
    [InlineData("conditional", false, 7.08D)]
    [InlineData("override", false, 0D)]
    [InlineData("override", true, 0D)]
    public void SaveAsPdf_AlignmentInheritsIndependentlyOfSpacing(string route, bool nativeDoc, double difference) {
        using var pdf = PdfPigDocument.Open(CreateLineAlignmentPdf(route, "top", nativeDoc, route is "table" or "conditional"));
        AssertLineAlignmentDifference(pdf, difference);
    }

    private static void AssertLineAlignmentDifference(PdfPigDocument pdf, double difference) {
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, x => x.Value == "A");
        var second = Assert.Single(letters, x => x.Value == "B");
        Assert.InRange(first.StartBaseLine.Y - second.StartBaseLine.Y, difference - .002D, difference + .002D);
    }

    private static byte[] CreateLineAlignmentPdf(string route, string alignment, bool nativeDoc, bool table,
        WordLineSpacingRule lineSpacingRule = WordLineSpacingRule.Exact) {
        const string family = "OfficeIMO device alignment";
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        var value = new VerticalTextAlignmentValues(alignment);
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
            new RunFonts { Ascii = family, HighAnsi = family }, new FontSize { Val = "16" })),
            new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(new SpacingBetweenLines {
                Line = "800", LineRule = LineSpacingRuleValues.Exact, Before = "0", After = "0"
            }, route == "document" ? new TextAlignment { Val = value } : null)));
        var cellTable = table ? source.AddTable(1, 1, WordTableStyle.TableGrid) : null;
        WordParagraph paragraph = cellTable == null ? source.AddParagraph() : cellTable.Rows[0].Cells[0].Paragraphs[0];
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(new Run(new RunProperties(new RunFonts { Ascii = family, HighAnsi = family },
            new FontSize { Val = "16" }), new Text("A")),
            new Run(new RunProperties(new RunFonts { Ascii = family, HighAnsi = family },
                new FontSize { Val = "32" }), new Text("B")));
        var properties = paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        if (route is "style" or "override") {
            styles.Append(new Style(new StyleName { Val = "Font alignment base" },
                new StyleParagraphProperties(new TextAlignment { Val = value })) {
                Type = StyleValues.Paragraph, StyleId = "FontAlignmentBase", CustomStyle = true
            }, new Style(new StyleName { Val = "Font alignment child" },
                new BasedOn { Val = "FontAlignmentBase" }) {
                Type = StyleValues.Paragraph, StyleId = "FontAlignmentChild", CustomStyle = true
            });
            properties.ParagraphStyleId = new ParagraphStyleId { Val = "FontAlignmentChild" };
        }
        if (route is "table" or "conditional") {
            var tableStyle = new Style(new StyleName { Val = "Table font alignment" }) {
                Type = StyleValues.Table, StyleId = "TableFontAlignment", CustomStyle = true
            };
            var tableProperties = new StyleParagraphProperties(new TextAlignment { Val = value });
            if (route == "table") tableStyle.Append(tableProperties);
            else tableStyle.Append(new TableStyleProperties(tableProperties) { Type = TableStyleOverrideValues.FirstRow });
            styles.Append(tableStyle);
            cellTable!._tableProperties!.TableStyle = new TableStyle { Val = "TableFontAlignment" };
            cellTable.ConditionalFormattingFirstRow = true;
        }
        if (route == "direct") properties.TextAlignment = new TextAlignment { Val = value };
        else if (route == "override") properties.TextAlignment = new TextAlignment { Val = VerticalTextAlignmentValues.Baseline };
        // A direct spacing declaration must not erase alignment inherited from elsewhere.
        paragraph.LineSpacing = (lineSpacingRule) switch { global::OfficeIMO.Word.WordLineSpacingRule.Exact => global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(40)), global::OfficeIMO.Word.WordLineSpacingRule.AtLeast => global::OfficeIMO.Word.WordLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Points(40)), _ => global::OfficeIMO.Word.WordLineSpacing.Multiple((40) / 12D) };
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }) : source.ToBytes()));
        var options = new PdfCreateOptions();
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics()));
        return document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }
}
