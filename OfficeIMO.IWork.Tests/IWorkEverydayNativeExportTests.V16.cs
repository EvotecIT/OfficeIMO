using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.IWork;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkEverydayNativeExportTests {
    private const string Apple16 = "native-exports/everyday-v16/";

    [Fact]
    public void Pages_16_keeps_the_full_two_page_body_in_editable_DOCX() {
        VerifyManifest("everyday-v16/field-notes.json");
        var source = IWorkSourceDocument.Load(Corpus(Apple16 + "field-notes.pages"));
        Assert.Throws<InvalidDataException>(() => source.ToWordResult());
        using var result = source.ToWordResult(new() { AllowPartialEditableReconstruction = true });
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_PAGES_DOCX_PRECISION");
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using var actual = WordprocessingDocument.Open(saved, false);
        using var apple = WordprocessingDocument.Open(Corpus(Apple16 + "apple-field-notes.docx"), false);

        string[] expected = Paragraphs(apple);
        Assert.Equal(67, expected.Length);
        Assert.Equal("End of inventory.", expected[^1]);
        Assert.Equal(expected, Paragraphs(actual));

        static string[] Paragraphs(WordprocessingDocument document) =>
            document.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Select(Text)
                .Where(text => text.Length != 0).ToArray();
    }

    [Fact]
    public void Keynote_16_keeps_ordered_slide_text_and_speaker_notes_in_editable_PPTX() {
        VerifyManifest("everyday-v16/briefing.json");
        var source = IWorkSourceDocument.Load(Corpus(Apple16 + "briefing.key"));
        Assert.Throws<InvalidDataException>(() => source.ToPowerPointResult());
        using var result = source.ToPowerPointResult(new() { AllowPartialEditableReconstruction = true });
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED");
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using var actual = PresentationDocument.Open(saved, false);
        using var apple = PresentationDocument.Open(Corpus(Apple16 + "apple-briefing.pptx"), false);

        (string Text, string Notes)[] expected = Slides(apple);
        Assert.Equal(new[] {
            ("Inventory briefing", "Speaker cue 1"),
            ("Count the stock", "Speaker cue 2"),
            ("Record the totals", "Speaker cue 3")
        }, expected);
        Assert.Equal(expected, Slides(actual));

        static (string, string)[] Slides(PresentationDocument document) {
            var presentation = document.PresentationPart!;
            return presentation.Presentation!.SlideIdList!.Elements<P.SlideId>()
                .Select(id => (SlidePart)presentation.GetPartById(id.RelationshipId!))
                .Select(slide => (
                    string.Concat(slide.Slide!.Descendants<A.Text>().Select(text => text.Text)),
                    string.Concat(slide.NotesSlidePart!.NotesSlide!.Descendants<P.Shape>()
                        .Where(shape => shape.Descendants<P.PlaceholderShape>().Any(placeholder =>
                            placeholder.Type?.Value == P.PlaceholderValues.Body))
                        .SelectMany(shape => shape.Descendants<A.Text>()).Select(text => text.Text))))
                .ToArray();
        }
    }

    [Fact]
    public void Numbers_16_keeps_ordered_tables_merges_and_cross_table_formula_caches_in_XLSX() {
        VerifyManifest("everyday-v16/inventory.json");
        var source = IWorkSourceDocument.Load(Corpus(Apple16 + "inventory.numbers"));
        Assert.Throws<InvalidDataException>(() => source.ToExcelResult());
        using var result = source.ToExcelResult(new() {
            Mode = IWorkConversionMode.EditableOnly,
            AllowPartialEditableReconstruction = true
        });
        IWorkTestPolicy.AssertCellPaddingOmission(result.Report);
        Assert.Empty(result.Report.SourceDeclarationIssues);
        Assert.Equal(new[] { (4, 3), (3, 2) }, result.Projection.Sheets.Single().Tables
            .Select(table => (table.RowCount, table.ColumnCount)));
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using var actual = SpreadsheetDocument.Open(saved, false);
        using var apple = SpreadsheetDocument.Open(Corpus(Apple16 + "apple-inventory.xlsx"), false);

        WorkbookPart reference = apple.WorkbookPart!;
        WorkbookPart output = actual.WorkbookPart!;
        var orderedNames = new[] { "Inventory - Inputs", "Inventory - Totals" };
        Assert.Equal(orderedNames, SheetNames(reference));
        Assert.Equal(orderedNames, SheetNames(output));

        Worksheet nativeInputs = SheetByName(reference, orderedNames[0]);
        Worksheet inputs = SheetByName(output, orderedNames[0]);
        // Numbers adds a merged title at row 1 and offsets the original grid by one row.
        Assert.Equal(new[] { "A1:C1", "B5:C5" }, MergeAddresses(nativeInputs));
        Assert.Equal(new[] { "B4:C4" }, MergeAddresses(inputs));
        Worksheet nativeTotals = SheetByName(reference, orderedNames[1]);
        Worksheet totals = SheetByName(output, orderedNames[1]);
        Assert.Equal(Data(reference, nativeInputs, 1), Data(output, inputs, 0));
        Assert.Equal(Data(reference, nativeTotals, 1), Data(output, totals, 0));
        Assert.Equal(new[] {
            ("B3", "SUM('Inventory - Inputs'!B3:B4)"),
            ("B4", "'Inventory - Inputs'!B3*'Inventory - Inputs'!C3+'Inventory - Inputs'!B4*'Inventory - Inputs'!C4")
        }, Formulas(nativeTotals));
        Assert.Equal(new[] {
            ("B2", "SUM('Inventory - Inputs'!B2:B3)"),
            ("B3", "'Inventory - Inputs'!B2*'Inventory - Inputs'!C2+'Inventory - Inputs'!B3*'Inventory - Inputs'!C3")
        }, Formulas(totals));

        static IEnumerable<string?> SheetNames(WorkbookPart workbook) =>
            workbook.Workbook!.Sheets!.Elements<Sheet>().Select(sheet => sheet.Name?.Value);
        static Worksheet SheetByName(WorkbookPart workbook, string name) {
            Sheet sheet = workbook.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name?.Value == name);
            return ((WorksheetPart)workbook.GetPartById(sheet.Id!)).Worksheet!;
        }
        static IEnumerable<string?> MergeAddresses(Worksheet worksheet) =>
            worksheet.Descendants<MergeCell>().Select(merge => merge.Reference?.Value);
        static IEnumerable<(string, string)> Formulas(Worksheet worksheet) =>
            worksheet.Descendants<Cell>().Where(cell => cell.CellFormula != null)
                .Select(cell => (cell.CellReference!.Value!, cell.CellFormula!.Text));
        static IEnumerable<(string Column, uint Row, string Kind, string Value)> Data(
            WorkbookPart workbook, Worksheet worksheet, uint titleRows) =>
            worksheet.Descendants<Row>().Where(row => row.RowIndex!.Value > titleRows)
                .SelectMany(row => row.Elements<Cell>().Select(cell => (
                    Column: string.Concat(cell.CellReference!.Value!.TakeWhile(char.IsLetter)),
                    Row: row.RowIndex!.Value - titleRows,
                    Content: Value(workbook, cell))))
                .Where(cell => cell.Content.value is not null)
                .Select(cell => (cell.Column, cell.Row, cell.Content.kind, cell.Content.value!));
    }
}
