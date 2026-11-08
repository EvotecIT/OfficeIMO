using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void ApplyingThemePreservesInheritedSemanticFontsAndEffectiveStyleToggles() {
        using var source = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true)) {
            MainDocumentPart main = package.AddMainDocumentPart();
            var styles = new Styles(new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }, new Bold(), new BoldComplexScript()))));
            styles.Append(new Style(new StyleName { Val = "Normal" }) { StyleId = "Normal", Type = StyleValues.Paragraph, Default = true });
            styles.Append(InheritedThemeStyle("ParagraphBold", StyleValues.Paragraph, null, new Bold(), new BoldComplexScript()));
            styles.Append(InheritedThemeStyle("ParagraphOff", StyleValues.Paragraph, null, new Bold { Val = false }, new BoldComplexScript { Val = false }));
            styles.Append(InheritedThemeStyle("CharacterBase", StyleValues.Character, null, new Bold(), new BoldComplexScript()));
            styles.Append(InheritedThemeStyle("CharacterChild", StyleValues.Character, "CharacterBase", new Bold(), new BoldComplexScript()));
            styles.Append(InheritedThemeStyle("CharacterFalse", StyleValues.Character, "CharacterBase", new Bold { Val = false }, new BoldComplexScript { Val = false }));
            styles.Append(InheritedThemeStyle("ParagraphBase", StyleValues.Paragraph, null, new RunFonts { Ascii = "Consolas" }, new Bold(), new BoldComplexScript()));
            styles.Append(InheritedThemeStyle("ParagraphChild", StyleValues.Paragraph, "ParagraphBase", new Bold(), new BoldComplexScript()));
            Style tableBase = InheritedThemeStyle("TableBase", StyleValues.Table, null, new RunFonts { Ascii = "Consolas" }, new Bold(), new BoldComplexScript());
            tableBase.Append(new TableStyleProperties(new RunPropertiesBaseStyle(new RunFonts { Ascii = "Symbol" }, new Bold(), new BoldComplexScript())) { Type = TableStyleOverrideValues.FirstRow });
            styles.Append(tableBase);
            Style tableChild = InheritedThemeStyle("TableChild", StyleValues.Table, "TableBase", new Bold(), new BoldComplexScript());
            tableChild.Append(new TableStyleProperties(new RunPropertiesBaseStyle(new Bold(), new BoldComplexScript())) { Type = TableStyleOverrideValues.FirstRow });
            styles.Append(tableChild);
            styles.Append(InheritedThemeStyle("TableInherited", StyleValues.Table, "TableBase"));
            Style cellOnly = InheritedThemeStyle("TableCellOnly", StyleValues.Table, "TableBase", new Bold(), new BoldComplexScript());
            cellOnly.Append(new TableStyleProperties(new TableStyleConditionalFormattingTableCellProperties(
                new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "EEEEEE" })) { Type = TableStyleOverrideValues.FirstRow });
            styles.Append(cellOnly);
            styles.Append(InheritedThemeStyle("NormalTable", StyleValues.Table, null, new RunFonts { Ascii = "Wingdings" }, new Bold(), new BoldComplexScript()));
            styles.Append(InheritedThemeStyle("TableNormal", StyleValues.Table, "NormalTable"));
            Style cornerBase = InheritedThemeStyle("CornerBase", StyleValues.Table, null, new RunFonts { Ascii = "Arial" }, new Bold(), new BoldComplexScript());
            cornerBase.Append(new TableStyleProperties(new RunPropertiesBaseStyle(new RunFonts { Ascii = "Consolas" })) { Type = TableStyleOverrideValues.WholeTable });
            cornerBase.Append(new TableStyleProperties(new RunPropertiesBaseStyle(new RunFonts { Ascii = "Courier New" }, new Bold { Val = false }, new BoldComplexScript { Val = false })) { Type = TableStyleOverrideValues.FirstRow });
            foreach (var corner in new[] {
                (TableStyleOverrideValues.NorthWestCell, "Symbol"), (TableStyleOverrideValues.NorthEastCell, "Wingdings"),
                (TableStyleOverrideValues.SouthWestCell, "Courier New"), (TableStyleOverrideValues.SouthEastCell, "Consolas")
            }) {
                cornerBase.Append(new TableStyleProperties(new RunPropertiesBaseStyle(new RunFonts { Ascii = corner.Item2 }, new Bold(), new BoldComplexScript())) { Type = corner.Item1 });
            }
            styles.Append(cornerBase);
            Style cornerChild = InheritedThemeStyle("CornerChild", StyleValues.Table, "CornerBase", new Bold(), new BoldComplexScript());
            cornerChild.Append(new TableStyleProperties(new RunPropertiesBaseStyle(new RunFonts { Ascii = "Symbol" }, new Bold { Val = false }, new BoldComplexScript { Val = false })) { Type = TableStyleOverrideValues.NorthWestCell });
            styles.Append(cornerChild);
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings(new Compatibility(new CompatibilitySetting {
                Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15"
            }));
            main.AddNewPart<FontTablePart>().Fonts = new Fonts(
                InheritedThemeFont("Symbol", "02", FontFamilyValues.Roman, FontPitchValues.Variable),
                InheritedThemeFont("Wingdings", "02", FontFamilyValues.Decorative, FontPitchValues.Variable),
                InheritedThemeFont("Consolas", "00", FontFamilyValues.Modern, FontPitchValues.Fixed),
                InheritedThemeFont("Courier New", "00", FontFamilyValues.Modern, FontPitchValues.Fixed),
                InheritedThemeFont("Arial", "00", FontFamilyValues.Swiss, FontPitchValues.Variable));
            main.Document = new Document(new Body(
                InheritedThemeParagraph("Default code"),
                InheritedThemeParagraph("Character cancellation", "ParagraphBold", "CharacterChild"),
                InheritedThemeParagraph("False style is a no-op", "ParagraphBold", "CharacterFalse"),
                InheritedThemeParagraph("Paragraph hierarchy reset", "ParagraphChild"),
                InheritedThemeParagraph("Paragraph resets default off", "ParagraphOff"),
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "ParagraphBold" }),
                    new Run(new RunProperties(new RunStyle { Val = "CharacterChild" }, new Bold { Val = false }, new BoldComplexScript { Val = false }), new Text("Direct off"))),
                InheritedThemeTable("TableChild", true, "Conditional symbols", "Whole-table code"),
                InheritedThemeParagraph("Normal-table control"),
                InheritedThemeTable("TableNormal", false, "Normal-table symbols"),
                InheritedThemeParagraph("Inherited-region control"),
                InheritedThemeTable("TableInherited", true, "Inherited conditional symbols"),
                InheritedThemeParagraph("Cell-only region control"),
                InheritedThemeTable("TableCellOnly", true, "Cell-only conditional symbols", "Cell-only whole-table code"),
                InheritedThemeParagraph("Active corner control"),
                InheritedThemeCornerTable("Active", new TableLook { Val = "07E0" }),
                InheritedThemeParagraph("Row-edge control"),
                InheritedThemeCornerTable("Rows only", new TableLook { Val = "0660" }),
                InheritedThemeParagraph("Column-edge control"),
                InheritedThemeCornerTable("Columns only", new TableLook { Val = "0780" }),
                InheritedThemeParagraph("Omitted-look control"),
                InheritedThemeCornerTable("Look omitted", null),
                InheritedThemeParagraph("Empty-look control"),
                InheritedThemeCornerTable("Look empty", new TableLook()),
                InheritedThemeParagraph("Expanded-look defaults control"),
                InheritedThemeCornerTable("Expanded defaults", new TableLook { Val = "0000", FirstRow = true })));
            AssertThemeInheritancePackageValid(package);
        }
        byte[] original = source.ToArray();
        source.Position = 0;
        using WordDocument document = WordDocument.Load(source);
        string authoredSettings = document.MainDocumentPartRoot.DocumentSettingsPart!.Settings!.OuterXml;
        string authoredFonts = document.MainDocumentPartRoot.FontTablePart!.Fonts!.OuterXml;
        string[] authoredDefinitions = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Where(style => style.StyleId != "Normal").Select(style => style.OuterXml).ToArray();
        document.ApplyTheme(ThemeForWordProof());
        string once = document._document.OuterXml;
        document.ApplyTheme(ThemeForWordProof());
        Assert.Equal(once, document._document.OuterXml);
        document.ApplyTheme(OfficeTheme.Dark);
        byte[] bytes = document.ToBytes();
        File.WriteAllBytes(Path.Combine(_directoryWithFiles, "ApplyTheme-inherited-before.docx"), original);
        File.WriteAllBytes(Path.Combine(_directoryWithFiles, "ApplyTheme-inherited-after.docx"), bytes);
        using var output = new MemoryStream(bytes, false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(output, false);
        AssertThemeInheritancePackageValid(saved);
        MainDocumentPart result = saved.MainDocumentPart!;
        Assert.Equal(authoredSettings, result.DocumentSettingsPart!.Settings!.OuterXml);
        Assert.Equal(authoredFonts, result.FontTablePart!.Fonts!.OuterXml);
        Assert.Equal(authoredDefinitions, result.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Where(style => style.StyleId != "Normal").Select(style => style.OuterXml));
        AssertInheritedThemeRun(result.Document, "Default code", "Courier New", true);
        AssertInheritedThemeRun(result.Document, "Character cancellation", "Courier New", true);
        AssertInheritedThemeRun(result.Document, "False style is a no-op", "Courier New", false);
        AssertInheritedThemeRun(result.Document, "Paragraph hierarchy reset", "Consolas", true);
        AssertInheritedThemeRun(result.Document, "Paragraph resets default off", "Courier New", false);
        AssertInheritedThemeRun(result.Document, "Direct off", "Courier New", false);
        AssertInheritedThemeRun(result.Document, "Conditional symbols", "Consolas", true);
        AssertInheritedThemeRun(result.Document, "Whole-table code", "Consolas", true);
        AssertInheritedThemeRun(result.Document, "Normal-table symbols", "Wingdings", true);
        AssertInheritedThemeRun(result.Document, "Inherited conditional symbols", "Symbol", true);
        AssertInheritedThemeRun(result.Document, "Cell-only conditional symbols", "Symbol", true);
        AssertInheritedThemeRun(result.Document, "Cell-only whole-table code", "Consolas", true);
        AssertInheritedThemeRun(result.Document, "Active NW", "Symbol", false);
        AssertInheritedThemeRun(result.Document, "Active NE", "Wingdings", true);
        AssertInheritedThemeRun(result.Document, "Active SW", "Courier New", true);
        AssertInheritedThemeRun(result.Document, "Active SE", "Consolas", true);
        AssertInheritedThemeRun(result.Document, "Rows only NW", "Courier New", false);
        AssertThemedBodyRun(result.Document, "Rows only SE");
        AssertThemedBodyRun(result.Document, "Columns only NW");
        AssertThemedBodyRun(result.Document, "Columns only SE");
        AssertInheritedThemeRun(result.Document, "Look omitted NW", "Symbol", false);
        AssertInheritedThemeRun(result.Document, "Look omitted NE", "Courier New", false);
        AssertThemedBodyRun(result.Document, "Look omitted SE");
        AssertInheritedThemeRun(result.Document, "Look empty NW", "Symbol", false);
        AssertInheritedThemeRun(result.Document, "Look empty NE", "Courier New", false);
        AssertInheritedThemeRun(result.Document, "Expanded defaults NW", "Symbol", false);
        AssertInheritedThemeRun(result.Document, "Expanded defaults NE", "Wingdings", true);
        AssertInheritedThemeRun(result.Document, "Expanded defaults SW", "Courier New", true);
        AssertInheritedThemeRun(result.Document, "Expanded defaults SE", "Consolas", true);
        TableProperties Look(string text) => ThemeRun(result.Document, text).Ancestors<Table>().First().GetFirstChild<TableProperties>()!;
        Assert.Null(Look("Look omitted NW").TableLook);
        Assert.Empty(Look("Look empty NW").TableLook!.GetAttributes());
        Assert.Equal("0000", Look("Expanded defaults NW").TableLook!.Val!.Value);
        Assert.True(Look("Expanded defaults NW").TableLook!.FirstRow!.Value);
    }

    private static Style InheritedThemeStyle(string id, StyleValues type, string? parent, params OpenXmlElement[] properties) {
        var style = new Style(new StyleName { Val = id }) { StyleId = id, Type = type, CustomStyle = true };
        if (parent != null) style.Append(new BasedOn { Val = parent });
        style.Append(new StyleRunProperties(properties));
        return style;
    }

    private static DocumentFormat.OpenXml.Wordprocessing.Font InheritedThemeFont(string name, string charset, FontFamilyValues family, FontPitchValues pitch) =>
        new(new FontCharSet { Val = charset }, new FontFamily { Val = family }, new Pitch { Val = pitch }) { Name = name };

    private static Paragraph InheritedThemeParagraph(string text, string? paragraphStyle = null, string? characterStyle = null) {
        var paragraph = new Paragraph();
        if (paragraphStyle != null) paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = paragraphStyle });
        var run = new Run(new Text(text));
        if (characterStyle != null) run.RunProperties = new RunProperties(new RunStyle { Val = characterStyle });
        paragraph.Append(run);
        return paragraph;
    }

    private static Table InheritedThemeTable(string style, bool firstRow, params string[] rows) {
        var table = new Table(
            new TableProperties(new TableStyle { Val = style }, new TableLook { Val = firstRow ? "0620" : "0600" }),
            new TableGrid(new GridColumn { Width = "4000" }));
        table.Append(rows.Select(text => new TableRow(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "4000" }),
            InheritedThemeParagraph(text)))));
        return table;
    }

    private static Table InheritedThemeCornerTable(string prefix, TableLook? look) {
        var properties = new TableProperties(new TableStyle { Val = "CornerChild" });
        if (look != null) properties.Append(look);
        return new Table(
            properties,
            new TableGrid(new GridColumn { Width = "1600" }, new GridColumn { Width = "1600" }, new GridColumn { Width = "1600" }),
            new TableRow(Cell("NW", 2), Cell("NE", 1)),
            new TableRow(Cell("SW", 1), Cell("SE", 2)));

        TableCell Cell(string corner, int span) => new(
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = (1600 * span).ToString(System.Globalization.CultureInfo.InvariantCulture) }, new GridSpan { Val = span }),
            InheritedThemeParagraph(prefix + " " + corner));
    }

    private static void AssertThemeInheritancePackageValid(WordprocessingDocument package) {
        ValidationErrorInfo[] errors = new OpenXmlValidator(FileFormatVersions.Office2010).Validate(package).ToArray();
        Assert.True(errors.Length == 0, FormatValidationErrors(errors));
    }

    private static void AssertThemedBodyRun(Document document, string text) {
        RunProperties properties = ThemeRun(document, text).RunProperties!;
        Assert.Null(properties.RunFonts!.Ascii);
        Assert.Equal(ThemeFontValues.MinorAscii, properties.RunFonts.AsciiTheme!.Value);
        Assert.True(properties.Bold!.Val?.Value ?? true);
        Assert.True(properties.BoldComplexScript!.Val?.Value ?? true);
    }

    private static void AssertInheritedThemeRun(Document document, string text, string font, bool bold) {
        RunProperties properties = ThemeRun(document, text).RunProperties!;
        Assert.Equal(font, properties.RunFonts!.Ascii!.Value);
        Assert.Null(properties.RunFonts.AsciiTheme);
        Assert.Equal(bold, properties.Bold!.Val?.Value ?? true);
        Assert.Equal(bold, properties.BoldComplexScript!.Val?.Value ?? true);
    }
}
