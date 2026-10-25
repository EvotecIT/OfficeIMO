using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_NoteTableRetainsSourceOrderCellStylesListsImagesAndLinks(bool endnote) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        note.RemoveAllChildren();
        WordList list = source.AddList(WordListStyle.Numbered);
        WordParagraph[] items = new[] { "NOTEBEFORE", "CELLITEMONE", "CELLITEMTWO", "NOTEAFTER" }
            .Select(text => list.AddItem(text)).ToArray();
        foreach (WordParagraph item in items) item._paragraph.Remove();
        items[0]._paragraph.InsertAfter(NoteBodyMarker(endnote), items[0]._paragraph.ParagraphProperties!);
        W.Table table = CreateNoteTable(new[] {
            new[] { "CELLHEADER", "OPPOSITEHEADER" }, new[] { "", "IMAGELEFT" }
        }, repeatHeader: true);
        table.GetFirstChild<W.TableProperties>()!.TableStyle = new W.TableStyle { Val = "NoteCellTypography" };
        table.GetFirstChild<W.TableProperties>()!.TableLook = new W.TableLook { FirstRow = true };
        W.TableCell[] cells = table.Elements<W.TableRow>().Last().Elements<W.TableCell>().ToArray();
        cells[0].RemoveAllChildren<W.Paragraph>();
        cells[0].Append(items[1]._paragraph, items[2]._paragraph);
        note.Append(items[0]._paragraph, table, items[3]._paragraph);
        items[3].AddBookmark("NoteTableAnchor");
        var imageParagraph = new WordParagraph(source, cells[1].Elements<W.Paragraph>().Single());
        using (var pixels = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng((byte)0, (byte)180, (byte)80)))
            imageParagraph.AddText(string.Empty).InsertImage(pixels, "note-table.png", 24, 16);
        var linkParagraph = new W.Paragraph();
        cells[1].Append(linkParagraph);
        new WordParagraph(source, linkParagraph).AddHyperLink("CELLLINK", new Uri("https://example.test/note-table"));
        cells[1].Append(new W.Paragraph(new W.Hyperlink(RichNoteRun("CELLANCHORLINK")) { Anchor = "NoteTableAnchor" }));
        source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style(
            new W.StyleName { Val = "Note cell typography" },
            new W.TableStyleProperties(new W.RunPropertiesBaseStyle(
                new W.RunFonts { Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName },
                new W.Color { Val = "803040" }, new W.FontSize { Val = "26" })) { Type = W.TableStyleOverrideValues.FirstRow }) {
                Type = W.StyleValues.Table, StyleId = "NoteCellTypography" });
        Assert.Empty(source.ValidateDocument());

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement loadedNotes = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        string original = main.Document.OuterXml + loadedNotes.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml;
        WordToPdfOptions options = FootnoteDiagnosticOptions();
        options.PdfOptions = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            string text = Regex.Replace(string.Concat(pdf.GetPages().Select(page => page.Text)), @"\s+", string.Empty);
            foreach (string token in new[] { "NOTEBEFORE", "CELLHEADER", "OPPOSITEHEADER", "CELLITEMONE", "CELLITEMTWO", "IMAGELEFT", "CELLLINK", "CELLANCHORLINK", "NOTEAFTER" })
                Assert.Single(Regex.Matches(text, token));
            foreach (string item in new[] { endnote ? "1.iNOTEBEFORE" : "1.1NOTEBEFORE", "2.CELLITEMONE", "3.CELLITEMTWO", "4.NOTEAFTER" })
                Assert.Contains(item, text, StringComparison.Ordinal);
            Assert.True(text.IndexOf("NOTEBEFORE", StringComparison.Ordinal) < text.IndexOf("CELLHEADER", StringComparison.Ordinal));
            Assert.True(text.IndexOf("CELLITEMTWO", StringComparison.Ordinal) < text.IndexOf("NOTEAFTER", StringComparison.Ordinal));
            var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
            var left = FindRichNoteToken(letters, "CELLHEADER");
            var right = FindRichNoteToken(letters, "OPPOSITEHEADER");
            Assert.True(right[0].StartBaseLine.X > left[left.Length - 1].EndBaseLine.X);
            Assert.Equal(left[0].StartBaseLine.Y, right[0].StartBaseLine.Y, 3);
            Assert.All(left.Concat(right), letter => {
                Assert.Contains("OfficeIMO", letter.FontName, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(13D, letter.PointSize, 3);
                Assert.Equal(128D / 255D, letter.Color.ToRGBValues().Item1, 3);
            });
            PdfImagePlacement image = Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
            Assert.Equal(18D, image.Width, 3);
            Assert.Equal(12D, image.Height, 3);
            PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
            Assert.Contains(navigation.LinkAnnotations, link => link.Uri == "https://example.test/note-table");
            Assert.Contains(navigation.LinkAnnotations, link => link.DestinationName == "NoteTableAnchor");
            Assert.Single(navigation.NamedDestinations, destination => destination.Name == "NoteTableAnchor");
            Assert.Single(navigation.NamedDestinations, destination => destination.Name.StartsWith(
                endnote ? "OfficeIMO.Endnote." : "OfficeIMO.Footnote.", StringComparison.Ordinal));
            Assert.DoesNotContain(result.Warnings, warning => warning.Code is
                "NativeFootnotePlacementApproximated" or "NativeEndnoteContentApproximated" or "NativeNestedTableLayoutApproximated");
            Assert.Equal(original, main.Document.OuterXml + loadedNotes.OuterXml + main.StyleDefinitionsPart.Styles.OuterXml);
        }
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_LoadedNoteTableWithOmittedOptionalPropertiesPreservesSourceXml(bool endnote) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        note.RemoveAllChildren();
        // Row and cell properties are optional; the table retains its required
        // empty tblPr. Conversion must not provision editing defaults to read it.
        note.Append(new W.Paragraph(NoteBodyMarker(endnote), RichNoteRun("BEFOREBARETABLE")),
            new W.Table(new W.TableProperties(),
                new W.TableGrid(new W.GridColumn { Width = "1800" }, new W.GridColumn { Width = "1800" }),
                new W.TableRow(new W.TableCell(new W.Paragraph(RichNoteRun("BARELEFT"))),
                    new W.TableCell(new W.Paragraph(RichNoteRun("BARERIGHT"))))),
            new W.Paragraph(RichNoteRun("AFTERBARETABLE")));
        // An explicit false override is valid note flow, not an authored page break.
        note.GetFirstChild<W.Paragraph>()!.ParagraphProperties =
            new W.ParagraphProperties(new W.PageBreakBefore { Val = false });
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement loadedNotes = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        W.Table loadedTable = Assert.Single(loadedNotes.Descendants<W.Table>());
        W.TableProperties loadedProperties = Assert.IsType<W.TableProperties>(loadedTable.GetFirstChild<W.TableProperties>());
        Assert.Empty(loadedProperties.ChildElements);
        Assert.All(loadedTable.Elements<W.TableRow>(), row => Assert.Null(row.TableRowProperties));
        Assert.All(loadedTable.Descendants<W.TableCell>(), cell => Assert.Null(cell.TableCellProperties));
        string original = loadedNotes.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        Assert.Equal(original, loadedNotes.OuterXml);
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
            var left = FindRichNoteToken(letters, "BARELEFT");
            var right = FindRichNoteToken(letters, "BARERIGHT");
            FindRichNoteToken(letters, "BEFOREBARETABLE");
            FindRichNoteToken(letters, "AFTERBARETABLE");
            Assert.Equal(left[0].StartBaseLine.Y, right[0].StartBaseLine.Y, 3);
            Assert.True(right[0].StartBaseLine.X > left[left.Length - 1].EndBaseLine.X);
            Assert.DoesNotContain(result.Warnings, warning => warning.Code is
                "NativeFootnotePlacementApproximated" or "NativeEndnoteContentApproximated");
            Assert.Equal(original, loadedNotes.OuterXml);
        }
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_NoteTableHeaderRetainsStylingWithoutRepeatingBodyPolicy(bool inNote, bool endnote) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        string[][] rows = new[] { new[] { "TABLEHEADLEFT", "TABLEHEADRIGHT" } }.Concat(
            Enumerable.Range(1, 18).Select(row => new[] { $"R{row:000}LEFT", $"R{row:000}RIGHT" })).ToArray();
        W.Table table = CreateNoteTable(rows, repeatHeader: true);
        foreach (W.Run run in table.Elements<W.TableRow>().First().Descendants<W.Run>())
            run.RunProperties = new W.RunProperties(new W.Bold(), new W.FontSize { Val = "20" });
        if (inNote) {
            OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
            note.RemoveAllChildren();
            note.Append(new W.Paragraph(NoteBodyMarker(endnote), RichNoteRun("NOTESTART")), table,
                new W.Paragraph(RichNoteRun("NOTEEND")));
        } else {
            source.AddParagraph("BODYSTART");
            source._wordprocessingDocument.MainDocumentPart!.Document.Body!.InsertBefore(table,
                source._wordprocessingDocument.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().Single());
        }
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string original = main.Document.OuterXml + main.FootnotesPart?.Footnotes?.OuterXml + main.EndnotesPart?.Endnotes?.OuterXml;
        WordToPdfOptions options = FootnoteDiagnosticOptions();
        options.PageSize = new PageSize(360D, 260D);
        options.Margins = PageMargins.Uniform(30D);
        options.PdfOptions = new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers };
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.True(pdf.NumberOfPages > 1);
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            int headerCount = Regex.Matches(text, "TABLEHEADLEFT").Count;
            if (inNote) Assert.Equal(1, headerCount);
            else Assert.True(headerCount > 1, "The ordinary body table retains its authored repeated header.");
            Assert.Equal(headerCount, Regex.Matches(text, "TABLEHEADRIGHT").Count);
            var firstPage = pdf.GetPages().First(page => page.Text.Contains("TABLEHEADLEFT", StringComparison.Ordinal));
            Assert.All(FindRichNoteToken(firstPage.Letters.ToArray(), "TABLEHEADLEFT"), letter => {
                Assert.Contains("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(10D, letter.PointSize, 3);
            });
            PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
            Assert.Contains(tagged.StructureElements, element => element.StructureType == "TH");
            foreach (string token in rows.Skip(1).SelectMany(row => row)) Assert.Single(Regex.Matches(text, token));
            Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeFootnotePlacementApproximated");
            Assert.Equal(original, main.Document.OuterXml + main.FootnotesPart?.Footnotes?.OuterXml + main.EndnotesPart?.Endnotes?.OuterXml);
        }
    }

    [Theory]
    [InlineData("positioned")]
    [InlineData("rotated")]
    public void SaveAsPdf_UnsupportedFootnoteTableProfileRetainsTextAndLossDiagnostic(string profile) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote: false);
        W.Table table = CreateNoteTable(new[] { new[] { "UNSUPPORTEDTABLETEXT", "OTHERCONTENT" } }, repeatHeader: false);
        if (profile == "positioned") table.GetFirstChild<W.TableProperties>()!.AddChild(new W.TablePositionProperties {
            HorizontalAnchor = W.HorizontalAnchorValues.Margin, VerticalAnchor = W.VerticalAnchorValues.Text,
            TablePositionX = 0, TablePositionY = 0 }, true);
        else table.Descendants<W.TableCell>().First().TableCellProperties!.Append(new W.TextDirection { Val = W.TextDirectionValues.TopToBottomRightToLeft });
        note.Append(table);
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var notes = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!;
        string original = notes.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("UNSUPPORTEDTABLETEXT", text, StringComparison.Ordinal);
        Assert.Single(result.Warnings, warning => warning.Code == "NativeFootnotePlacementApproximated");
        Assert.True(result.HasLoss);
        Assert.Equal(original, notes.OuterXml);
    }

    private static W.Table CreateNoteTable(IEnumerable<string[]> rows, bool repeatHeader) {
        var table = new W.Table(new W.TableProperties(new W.TableWidth { Type = W.TableWidthUnitValues.Pct, Width = "5000" },
            new W.TableBorders(new W.TopBorder { Val = W.BorderValues.Single, Size = 4U }, new W.LeftBorder { Val = W.BorderValues.Single, Size = 4U },
                new W.BottomBorder { Val = W.BorderValues.Single, Size = 4U }, new W.RightBorder { Val = W.BorderValues.Single, Size = 4U },
                new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 4U }, new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 4U }),
            new W.TableLayout { Type = W.TableLayoutValues.Fixed }),
            new W.TableGrid(new W.GridColumn { Width = "3000" }, new W.GridColumn { Width = "3000" }));
        foreach (string[] row in rows) {
            var nativeRow = new W.TableRow();
            if (repeatHeader && !table.Elements<W.TableRow>().Any()) nativeRow.Append(new W.TableRowProperties(new W.TableHeader()));
            foreach (string text in row) nativeRow.Append(new W.TableCell(new W.TableCellProperties(
                    new W.TableCellWidth { Type = W.TableWidthUnitValues.Dxa, Width = "3000" }),
                new W.Paragraph(new W.ParagraphProperties(new W.SpacingBetweenLines {
                    Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Exact }),
                    RichNoteRun(text))));
            table.Append(nativeRow);
        }
        return table;
    }
}
