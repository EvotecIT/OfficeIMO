using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ApplyingThemeIsExplicitAndPreservesLoadedContentAndGeometry(bool apply, bool dark) {
        byte[] source = CreateAuthoredThemeSource();
        using var input = new MemoryStream(source, writable: false);
        using WordDocument document = WordDocument.Load(input);
        MainDocumentPart before = document._wordprocessingDocument.MainDocumentPart!;
        string sourceStyles = before.StyleDefinitionsPart!.Styles!.OuterXml;
        string sourceBackground = before.Document.DocumentBackground!.OuterXml;
        string[] text = before.Document.Descendants<Text>().Select(item => item.Text).ToArray();
        string section = before.Document.Body!.GetFirstChild<SectionProperties>()!.OuterXml;
        string[] geometry = before.Document.Descendants<Table>().Select(ThemeTableGeometry).ToArray();
        string drawing = before.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Single().OuterXml;
        string[] fieldCodes = before.Document.Descendants<SimpleField>().Select(field => field.Instruction!.Value!).ToArray();
        string[] external = before.HyperlinkRelationships.Select(relationship => relationship.Id + "=" + relationship.Uri).ToArray();
        byte[][] pictures = before.ImageParts.Select(part => { using Stream stream = part.GetStream(); using var bytes = new MemoryStream(); stream.CopyTo(bytes); return bytes.ToArray(); }).ToArray();
        OfficeTheme theme = dark ? OfficeTheme.Dark : ThemeForWordProof();
        if (apply) Assert.Same(document, document.ApplyTheme(theme));

        byte[] savedBytes = document.ToBytes();
        File.WriteAllBytes(Path.Combine(_directoryWithFiles, "ApplyTheme-before.docx"), source);
        File.WriteAllBytes(Path.Combine(_directoryWithFiles, apply ? dark ? "ApplyTheme-dark-after.docx" : "ApplyTheme-custom-after.docx" : "ApplyTheme-faithful.docx"), savedBytes);
        using var output = new MemoryStream(savedBytes, writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(output, false);
        MainDocumentPart main = saved.MainDocumentPart!;
        Assert.Equal(text, main.Document.Descendants<Text>().Select(item => item.Text));
        Assert.Equal(section, main.Document.Body!.GetFirstChild<SectionProperties>()!.OuterXml);
        Assert.Equal(geometry, main.Document.Descendants<Table>().Select(ThemeTableGeometry));
        Assert.Equal(drawing, main.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Single().OuterXml);
        Assert.Equal(fieldCodes, main.Document.Descendants<SimpleField>().Select(field => field.Instruction!.Value!));
        Assert.Equal(external, main.HyperlinkRelationships.Select(relationship => relationship.Id + "=" + relationship.Uri));
        Assert.Equal("Target", Assert.Single(main.Document.Descendants<BookmarkStart>()).Name!.Value);
        Assert.Equal(pictures.Length, main.ImageParts.Count());
        foreach (var pair in pictures.Zip(main.ImageParts, (bytes, part) => (bytes, part))) {
            using Stream stream = pair.part.GetStream(); using var bytes = new MemoryStream(); stream.CopyTo(bytes);
            Assert.Equal(pair.bytes, bytes.ToArray());
        }
        Run body = ThemeRun(main.Document, "Authored body");
        Assert.NotNull(body.RunProperties!.Bold);
        Assert.NotNull(body.RunProperties.Italic);
        Assert.Equal("pl-PL", body.RunProperties.Languages!.Val!.Value);
        Assert.False(body.RunProperties.RightToLeftText!.Val!.Value);
        Assert.Equal("Courier New", ThemeRun(main.Document, "inline code").RunProperties!.RunFonts!.Ascii!.Value);
        Assert.Equal("Wingdings", main.Document.Descendants<SymbolChar>().Single().Font!.Value);
        if (apply) {
            DocumentBackground canvas = main.Document.DocumentBackground!;
            Assert.Equal("FFFFFF", canvas.Color!.Value);
            Assert.Null(canvas.ThemeColor);
            Assert.Null(canvas.ThemeTint);
            Assert.Null(canvas.ThemeShade);
            Assert.Equal(WordThemeDefaults.PageText(theme).ToRgbHex(), body.RunProperties.Color!.Val!.Value);
            Assert.Equal(ThemeFontValues.MinorAscii, body.RunProperties.RunFonts!.AsciiTheme!.Value);
            Assert.Equal("22", body.RunProperties.FontSize!.Val!.Value);
            Run heading = ThemeRun(main.Document, "Semantic heading");
            Assert.Equal(dark ? "25" : "36", heading.RunProperties!.FontSize!.Val!.Value);
            Assert.False(heading.RunProperties.Bold!.Val!.Value);
            Assert.Equal(dark ? "240" : "580", heading.Ancestors<Paragraph>().First().ParagraphProperties!.SpacingBetweenLines!.Before!.Value);
            Assert.Equal(dark ? "56" : "54", ThemeRun(main.Document, "Document title").RunProperties!.FontSize!.Val!.Value);
            Assert.Equal(dark ? "25" : "36", ThemeRun(main.Document, "Document subtitle").RunProperties!.FontSize!.Val!.Value);
            Assert.Equal("Consolas", ThemeRun(main.Document, "Code block").RunProperties!.RunFonts!.Ascii!.Value);
            Assert.Equal(OfficeTheme.Modern.Surfaces.Surface.Resolve(theme).ToRgbHex(), ThemeRun(main.Document, "Code block").Ancestors<Paragraph>().First().ParagraphProperties!.Shading!.Fill!.Value);
            Table[] tables = main.Document.Descendants<Table>().ToArray();
            Assert.Equal(dark ? "1E2A39" : "375A74", tables[0].Descendants<TableCell>().First().TableCellProperties!.Shading!.Fill!.Value);
            Assert.Equal("FFFFFF", tables[1].Descendants<TableCell>().First().TableCellProperties!.Shading!.Fill!.Value);
            Assert.NotNull(ThemeRun(main.Document, "Declared header").RunProperties!.Bold);
            Assert.Equal("22", ThemeRun(main.HeaderParts.Single().Header, "Header story").RunProperties!.FontSize!.Val!.Value);
            Assert.Equal("22", ThemeRun(main.FootnotesPart!.Footnotes!, "Note story").RunProperties!.FontSize!.Val!.Value);
            Assert.Equal(dark ? "68A1F9" : "2468AC", ThemeRun(main.Document, "Linked text").RunProperties!.Color!.Val!.Value);
            Assert.Equal(WordThemeDefaults.PageText(theme).ToRgbHex(), heading.RunProperties.Color!.Val!.Value);
        } else {
            Assert.Equal(sourceBackground, main.Document.DocumentBackground!.OuterXml);
            Assert.Equal(sourceStyles, main.StyleDefinitionsPart!.Styles!.OuterXml);
            Assert.Equal("AC2345", body.RunProperties.Color!.Val!.Value);
            Assert.Equal("Times New Roman", body.RunProperties.RunFonts!.Ascii!.Value);
        }
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(saved));
    }

    [Fact]
    public void ApplyingThemeToStylelessDocumentIsIdempotentAndSwitchesOwnedFontsAndBold() {
        using var input = CreateBareThemeSource();
        using WordDocument document = WordDocument.Load(input);
        MainDocumentPart main = document._wordprocessingDocument.MainDocumentPart!;
        Assert.Null(main.StyleDefinitionsPart);
        OfficeTheme first = ThemeForWordProof().WithFonts(new OfficeThemeFonts(new OfficeThemeFont("Courier New", Array.Empty<string>()), new OfficeThemeFont("Courier New", Array.Empty<string>())));
        document.ApplyTheme(first);
        Assert.Equal("Courier New", document.Settings.FontFamily);
        string firstXml = main.Document.OuterXml;
        string firstStyles = main.StyleDefinitionsPart!.Styles!.OuterXml;
        document.ApplyTheme(first);
        Assert.Equal(firstXml, main.Document.OuterXml);
        Assert.Equal(firstStyles, main.StyleDefinitionsPart!.Styles!.OuterXml);
        document.AddParagraph("Added after mono theme");
        OfficeTheme second = ThemeForWordProof().WithTextStyle(first.TextStyle.With(headingBold: false));
        document.ApplyTheme(second);
        Assert.Equal("Arial", document.Settings.FontFamily);
        Assert.Null(main.Document.DocumentBackground);
        Assert.Empty(main.StyleDefinitionsPart.Styles!.Elements<Style>());
        Assert.Equal("Arial", main.ThemePart!.Theme.ThemeElements!.FontScheme!.MinorFont!.LatinFont!.Typeface!.Value);
        Assert.Equal(ThemeFontValues.MinorAscii, ThemeRun(main.Document, "Plain body").RunProperties!.RunFonts!.AsciiTheme!.Value);
        RunFonts addedFonts = ThemeRun(main.Document, "Added after mono theme").RunProperties!.RunFonts!;
        Assert.Equal(ThemeFontValues.MinorAscii, addedFonts.AsciiTheme!.Value);
        Assert.Null(addedFonts.Ascii);
        Assert.False(ThemeRun(main.Document, "Outline heading").RunProperties!.Bold!.Val!.Value);
        Assert.Equal("Courier New", ThemeRun(main.Document, "authored code").RunProperties!.RunFonts!.Ascii!.Value);
        Assert.True(ThemeRun(main.Document, "authored emphasis").RunProperties!.Bold!.Val?.Value ?? true);
        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordprocessingDocument saved = WordprocessingDocument.Open(savedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Fact]
    public async Task NativeWordConversionThemesOnlyItsOperationOwnedModel() {
        string sourcePath = Path.Combine(_directoryWithFiles, "ApplyTheme.Native.Source.docx");
        string destination = Path.Combine(_directoryWithFiles, "ApplyTheme.Native.Themed.dotx");
        string faithful = Path.Combine(_directoryWithFiles, "ApplyTheme.Native.Faithful.dotx");
        using (MemoryStream source = CreateBareThemeSource()) File.WriteAllBytes(sourcePath, source.ToArray());
        byte[] original = File.ReadAllBytes(sourcePath);
        var options = new WordDocumentConversionOptions {
            Theme = ThemeForWordProof(),
            LoadOptions = new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.SaveOnDispose },
            FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace
        };
        WordDocumentConversionReport analysis = await WordDocument.AnalyzeConversionAsync(sourcePath, destination, options);
        Assert.Equal("Word.Docx", analysis.SourceFormatDescriptor.Id);
        Assert.Equal("Word.Dotx", analysis.DestinationFormatDescriptor.Id);
        Assert.Equal(original, File.ReadAllBytes(sourcePath));
        WordDocumentConversionResult themedResult = await WordDocument.ConvertAsync(sourcePath, destination, options);
        WordDocumentConversionResult faithfulResult = await WordDocument.ConvertAsync(sourcePath, faithful, new WordDocumentConversionOptions { FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace });
        themedResult.ThrowIfLoss();
        faithfulResult.ThrowIfLoss();
        Assert.Equal(original, File.ReadAllBytes(sourcePath));
        using WordDocument themed = WordDocument.Load(destination);
        using WordDocument unthemed = WordDocument.Load(faithful);
        Assert.Equal(WordprocessingDocumentType.Template, themed._wordprocessingDocument.DocumentType);
        Assert.Equal(WordprocessingDocumentType.Template, unthemed._wordprocessingDocument.DocumentType);
        Assert.Equal("223344", ThemeRun(themed._document, "Plain body").RunProperties!.Color!.Val!.Value);
        Assert.Equal("AC2345", ThemeRun(unthemed._document, "Plain body").RunProperties!.Color!.Val!.Value);
    }

    private static OfficeTheme ThemeForWordProof() => OfficeTheme.Modern
        .WithFonts(new OfficeThemeFonts(new OfficeThemeFont("Georgia", Array.Empty<string>()), new OfficeThemeFont("Arial", Array.Empty<string>())))
        .WithPalette(OfficeTheme.Modern.Palette.With(text: OfficeColor.ParseHex("223344")))
        .WithTypeScale(OfficeTheme.Modern.TypeScale.With(title: OfficeLength.Points(27), heading3: OfficeLength.Points(18), body: OfficeLength.Points(11)))
        .WithTextStyle(OfficeTheme.Modern.TextStyle.With(headingBold: true, linkColor: OfficeColor.ParseHex("2468AC"),
            heading3Spacing: new OfficeThemeParagraphSpacing(OfficeLength.Points(29), OfficeLength.Points(7))))
        .WithTableStyle(OfficeTheme.Modern.TableStyle.With(headerFill: OfficeColor.ParseHex("375A74"), headerText: OfficeColor.White, cellPadding: OfficeLength.Points(4)));

    private static Run ThemeRun(OpenXmlElement root, string text) => root.Descendants<Run>().Single(run => run.Elements<Text>().Any(item => item.Text == text));

    private static MemoryStream CreateBareThemeSource() {
        var stream = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true)) {
            package.AddMainDocumentPart().Document = new Document(new Body(
                new Paragraph(new Run(new RunProperties(new RunFonts { Ascii = "Times New Roman" }, new Color { Val = "AC2345" }), new Text("Plain body"))),
                new Paragraph(new ParagraphProperties(new OutlineLevel { Val = 2 }), new Run(new Text("Outline heading"))),
                new Paragraph(new Run(new RunProperties(new RunFonts { Ascii = "Courier New" }), new Text("authored code")),
                    new Run(new RunProperties(new Bold(), new Italic()), new Text("authored emphasis")))));
        }
        stream.Position = 0;
        return stream;
    }

    private static byte[] CreateAuthoredThemeSource() {
        using var input = CreateBareThemeSource();
        using (WordprocessingDocument package = WordprocessingDocument.Open(input, true)) {
            MainDocumentPart main = package.MainDocumentPart!;
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleName { Val = "Authored normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                new Style(new StyleName { Val = "Project heading" }, new StyleParagraphProperties(new OutlineLevel { Val = 2 })) { Type = StyleValues.Paragraph, StyleId = "ProjectHeading" },
                new Style(new StyleName { Val = "Document title" }) { Type = StyleValues.Paragraph, StyleId = "Title" },
                new Style(new StyleName { Val = "Document subtitle" }) { Type = StyleValues.Paragraph, StyleId = "Subtitle" },
                new Style(new StyleName { Val = "Code block" }, new StyleRunProperties(new RunFonts { Ascii = "Consolas" })) { Type = StyleValues.Paragraph, StyleId = "OfficeIMOCodeBlock" });
            main.Document.DocumentBackground = new DocumentBackground { Color = "101820", ThemeColor = ThemeColorValues.Dark1, ThemeTint = "20", ThemeShade = "80" };
            Body body = main.Document.Body!;
            body.RemoveAllChildren();
            body.Append(new Paragraph(new BookmarkStart { Name = "Target", Id = "1" },
                new Run(new RunProperties(new RunFonts { Ascii = "Times New Roman" }, new Bold(), new Italic(), new Color { Val = "AC2345" },
                    new RightToLeftText { Val = false }, new Languages { Val = "pl-PL" }), new Text("Authored body")), new BookmarkEnd { Id = "1" }));
            body.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "ProjectHeading" }), new Run(new RunProperties(new Bold { Val = false }), new Text("Semantic heading"))));
            body.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Title" }), new Run(new Text("Document title"))));
            body.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Subtitle" }), new Run(new Text("Document subtitle"))));
            body.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "OfficeIMOCodeBlock" }), new Run(new Text("Code block"))));
            body.Append(new Paragraph(new Run(new RunProperties(new RunFonts { Ascii = "Courier New" }), new Text("inline code")), new Run(new SymbolChar { Font = "Wingdings", Char = "F0FC" })));
            string linkId = main.AddHyperlinkRelationship(new Uri("https://example.com/source"), true).Id;
            body.Append(new Paragraph(new Hyperlink(new Run(new Text("Linked text"))) { Id = linkId },
                new SimpleField(new Run(new Text("1"))) { Instruction = " PAGE " }, new Run(new FootnoteReference { Id = 1 })));
            body.Append(CreateThemeSourceTable(true), CreateThemeSourceTable(false));
            HeaderPart header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(new Run(new Text("Header story"))));
            main.AddNewPart<FootnotesPart>().Footnotes = new Footnotes(new Footnote(new Paragraph(new Run(new Text("Note story")))) { Id = 1 });
            body.Append(new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) },
                new PageSize { Width = 12240U, Height = 15840U }, new PageMargin { Top = 720, Bottom = 1200, Left = 900U, Right = 1100U, Header = 400U, Footer = 400U, Gutter = 0U }));
        }
        input.Position = 0;
        using WordDocument document = WordDocument.Load(input);
        using Stream image = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "NativeOptionAssets", "OfficeIMO.png"));
        document.AddImage(image, OfficeLength.Points(24), OfficeLength.Points(24));
        return document.ToBytes();
    }

    private static Table CreateThemeSourceTable(bool header) => new(
        new TableProperties(new TableWidth { Type = TableWidthUnitValues.Dxa, Width = "5000" },
            new TableJustification { Val = TableRowAlignmentValues.Center }, new DocumentFormat.OpenXml.Wordprocessing.TableLayout { Type = TableLayoutValues.Fixed },
            new TableLook { FirstRow = false }),
        new TableGrid(new GridColumn { Width = "2400" }, new GridColumn { Width = "2600" }),
        new TableRow(new TableRowProperties(new CantSplit(), new TableRowHeight { Val = 600U, HeightType = HeightRuleValues.AtLeast }, new TableHeader { Val = header ? OnOffOnlyValues.On : OnOffOnlyValues.Off }),
            new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "5000" }, new GridSpan { Val = 2 }),
                new Paragraph(new Run(new Text(header ? "Declared header" : "Ordinary first row"))))),
        new TableRow(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2400" }), new Paragraph(new Run(new Text("Value")))),
            new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2600" }), new Paragraph(new Run(new Text("42"))))));

    private static string ThemeTableGeometry(Table table) {
        Table clone = (Table)table.CloneNode(true);
        foreach (OpenXmlElement paint in clone.Descendants().Where(element => element is Paragraph or Shading or TableBorders or TableCellBorders or TableCellMargin or TableCellMarginDefault).ToArray()) paint.Remove();
        return clone.OuterXml;
    }
}
