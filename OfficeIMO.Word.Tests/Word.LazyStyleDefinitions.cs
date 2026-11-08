using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class WordLoadDeferredInitializationTests {
    [Fact]
    public void CreatingTablesAddsOnlyTheirStyleAndItsDependencies() {
        using WordDocument document = WordDocument.Create();
        MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
        Assert.Equal(new[] { "TableNormal" }, GetTableStyleIds(mainPart));
        Assert.True(mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Single(style => style.StyleId == "TableNormal").Default!.Value);

        document.AddTable(1, 1, WordTableStyle.TableGrid);
        document.AddTable(1, 1, WordTableStyle.TableGrid);

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        Assert.Equal(new[] { "TableGrid", "TableNormal" }, GetTableStyleIds(saved.MainDocumentPart!));
        Assert.All(saved.MainDocumentPart!.Document.Body!.Elements<Table>(), table =>
            Assert.Equal("TableGrid", table.GetFirstChild<TableProperties>()!.TableStyle!.Val!.Value));
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AssigningStyleToUnstyledLoadedTableAddsItsDefinition(bool useStyleId) {
        using var input = CreateMinimalDocument(table: CreateUnstyledTable());
        using WordDocument document = WordDocument.Load(input);
        WordTable table = Assert.Single(document.Tables);
        Assert.Null(table.Style);

        if (useStyleId) table.SetStyleId("TableGrid");
        else table.Style = WordTableStyle.TableGrid;

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        Table persisted = Assert.Single(saved.MainDocumentPart!.Document.Body!.Elements<Table>());
        TableStyle style = Assert.IsType<TableStyle>(persisted.GetFirstChild<TableProperties>()?.TableStyle);
        Assert.Equal("TableGrid", style.Val!.Value);
        Assert.Equal(new[] { "Normal", "TableGrid", "TableNormal" }, GetStyleIds(saved.MainDocumentPart!).OrderBy(id => id, StringComparer.Ordinal));
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Fact]
    public void ApplyingBuiltInStylePreservesAnAuthoredDefinitionAndItsBase() {
        Style authoredBase = new(new StyleName { Val = "Project base" }) {
            Type = StyleValues.Table,
            StyleId = "ProjectBase"
        };
        Style authoredGrid = new(
            new StyleName { Val = "Project grid" },
            new BasedOn { Val = "ProjectBase" },
            new StyleTableProperties(new TableBorders(new TopBorder {
                Val = BorderValues.Double,
                Color = "B3261E",
                Size = 12U
            }))) {
            Type = StyleValues.Table,
            StyleId = "TableGrid"
        };
        using var input = CreateMinimalDocument(
            table: CreateUnstyledTable(), additionalStyles: new[] { authoredBase, authoredGrid });
        using WordDocument document = WordDocument.Load(input);
        Styles original = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        string originalXml = original.OuterXml;

        Assert.Single(document.Tables).Style = WordTableStyle.TableGrid;

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        Assert.Equal(originalXml, saved.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml);
        Assert.Equal(new[] { "ProjectBase", "TableGrid" }, GetTableStyleIds(saved.MainDocumentPart!));
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Fact]
    public void CustomTableStyleCanBeInspectedAndClearedWithoutRemovingItsDefinition() {
        Table inputTable = CreateUnstyledTable();
        inputTable.GetFirstChild<TableProperties>()!.TableStyle = new TableStyle { Val = "ProjectTable" };
        Style authored = new(new StyleName { Val = "Project table" }) {
            Type = StyleValues.Table,
            StyleId = "ProjectTable"
        };
        using var input = CreateMinimalDocument(table: inputTable, additionalStyles: new[] { authored });
        using WordDocument document = WordDocument.Load(input);

        WordTable table = Assert.Single(document.Tables);
        Assert.Null(table.Style);
        table.Style = null;

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        Table persisted = Assert.Single(saved.MainDocumentPart!.Document.Body!.Elements<Table>());
        Assert.Null(persisted.GetFirstChild<TableProperties>()?.TableStyle);
        Assert.Equal(new[] { "ProjectTable" }, GetTableStyleIds(saved.MainDocumentPart!));
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Fact]
    public void AddingStyledContentToStylelessDocumentResolvesItsStyleReferences() {
        using var input = CreateMinimalDocument(includeStyles: false);
        using WordDocument document = WordDocument.Load(input);
        document.AddList(WordListStyle.Numbered).AddItem("List entry");
        document.AddTableOfContent();
        WordFooter footer = document.PageSetup.Footer;
        footer.AddPageNumber(WordPageNumberStyle.PlainNumber);

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        MainDocumentPart mainPart = saved.MainDocumentPart!;
        var defined = new HashSet<string>(GetStyleIds(mainPart), StringComparer.Ordinal);
        Assert.Contains("ListParagraph", defined);
        Assert.Contains("TOCHeading", defined);
        Assert.Empty(GetTableStyleIds(mainPart));
        var roots = new[] { mainPart.Document }.Cast<DocumentFormat.OpenXml.OpenXmlElement>()
            .Concat(mainPart.FooterParts.Select(part => part.Footer));
        foreach (var root in roots) {
            Assert.All(root.Descendants<ParagraphStyleId>(), style => Assert.Contains(style.Val!.Value!, defined));
            Assert.All(root.Descendants<RunStyle>(), style => Assert.Contains(style.Val!.Value!, defined));
        }
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2013).Validate(saved));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AddingNoteToLoadedDocumentAddsOnlyTheNoteStyles(bool endnote) {
        using var input = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(input);
        WordParagraph paragraph = Assert.Single(document.Paragraphs);

        if (endnote) paragraph.AddEndNote("Source note");
        else paragraph.AddFootNote("Source note");

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        string prefix = endnote ? "Endnote" : "Footnote";
        string[] expected = { "DefaultParagraphFont", "Normal", prefix + "Reference", prefix + "Text", prefix + "TextChar" };
        Assert.Equal(expected.OrderBy(id => id, StringComparer.Ordinal), GetStyleIds(saved.MainDocumentPart!).OrderBy(id => id, StringComparer.Ordinal));
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Fact]
    public void AssigningHeadingToLoadedParagraphAddsItsLinkedStyleAndBase() {
        using var input = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(input);

        Assert.Single(document.Paragraphs).Style = WordParagraphStyles.Heading1;

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        Assert.Equal(new[] { "DefaultParagraphFont", "Heading1", "Heading1Char", "Normal" },
            GetStyleIds(saved.MainDocumentPart!).OrderBy(id => id, StringComparer.Ordinal));
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Theory]
    [InlineData("paragraph", true)]
    [InlineData("character", true)]
    [InlineData("table", true)]
    [InlineData("paragraph", false)]
    [InlineData("character", false)]
    [InlineData("table", false)]
    public void AddingBuiltInStylePreservesTheSourceDefaultForItsType(string kind, bool sourceHasDefault) {
        StyleValues type = kind switch {
            "paragraph" => StyleValues.Paragraph,
            "character" => StyleValues.Character,
            _ => StyleValues.Table
        };
        using var input = CreateMinimalDocument(includeStyles: false);
        using (WordprocessingDocument package = WordprocessingDocument.Open(input, true)) {
            StyleDefinitionsPart stylePart = package.MainDocumentPart!.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles = new Styles(new Style(
                new StyleName { Val = "Authored source style" },
                new StyleRunProperties(new Color { Val = "315F23" })) {
                Type = type,
                StyleId = "ProjectDefault",
                Default = sourceHasDefault
            });
        }
        input.Position = 0;
        using WordDocument document = WordDocument.Load(input);
        string authoredXml = document._wordprocessingDocument.MainDocumentPart!
            .StyleDefinitionsPart!.Styles!.Elements<Style>().Single().OuterXml;
        string selectedId;
        string baseId;
        if (kind == "paragraph") {
            Assert.Single(document.Paragraphs).Style = WordParagraphStyles.Heading1;
            selectedId = "Heading1";
            baseId = "Normal";
        } else if (kind == "character") {
            Assert.Single(document.Paragraphs).CharacterStyle = WordCharacterStyles.Heading1Char;
            selectedId = "Heading1Char";
            baseId = "DefaultParagraphFont";
        } else {
            document.AddTable(1, 1, WordTableStyle.TableGrid);
            selectedId = "TableGrid";
            baseId = "TableNormal";
        }

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        MainDocumentPart mainPart = saved.MainDocumentPart!;
        Style[] styles = mainPart.StyleDefinitionsPart!.Styles!.Elements<Style>().ToArray();
        Style[] defaults = styles.Where(style => style.Type?.Value == type && style.Default?.Value == true).ToArray();
        if (sourceHasDefault) Assert.Equal("ProjectDefault", Assert.Single(defaults).StyleId!.Value);
        else Assert.Empty(defaults);
        Assert.Equal(authoredXml, Assert.Single(styles, style => style.StyleId == "ProjectDefault").OuterXml);
        Assert.Contains(styles, style => style.StyleId?.Value == selectedId);
        Assert.Contains(styles, style => style.StyleId?.Value == baseId);
        Assert.Equal("Minimal document", mainPart.Document.Body!.InnerText);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Fact]
    public void AddingHeaderWatermarkToLoadedDocumentDefinesItsHeaderStyles() {
        using var input = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(input);

        document.PageSetup.Header.AddWatermark(WordWatermarkStyle.Text, "Draft");

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        MainDocumentPart mainPart = saved.MainDocumentPart!;
        var defined = new HashSet<string>(GetStyleIds(mainPart), StringComparer.Ordinal);
        Assert.Equal(new[] { "DefaultParagraphFont", "Header", "HeaderChar", "Normal" },
            defined.OrderBy(id => id, StringComparer.Ordinal));
        Header header = Assert.Single(mainPart.HeaderParts).Header;
        Assert.Contains(header.Descendants<ParagraphStyleId>(), style => style.Val?.Value == "Header");
        Assert.All(header.Descendants<ParagraphStyleId>(), style => Assert.Contains(style.Val!.Value!, defined));
        Assert.All(header.Descendants<RunStyle>(), style => Assert.Contains(style.Val!.Value!, defined));
        Assert.Equal("Minimal document", mainPart.Document.Body!.InnerText);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2013).Validate(saved));
    }

    private static Table CreateUnstyledTable() => new(
        new TableProperties(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto }),
        new TableGrid(new GridColumn { Width = "2400" }),
        new TableRow(new TableCell(
            new TableCellProperties(new TableCellWidth { Width = "2400", Type = TableWidthUnitValues.Dxa }),
            new Paragraph(new Run(new Text("Table text"))))));
}
