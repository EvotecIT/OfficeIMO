using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNumericFormatWritingTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Explicit_numeric_formats_retain_semantics_in_all_shared_native_table_codecs(IWorkDocumentKind kind) {
        var formats = Formats(); var data = IWorkTableData.Create(formats.Length, 2);
        for (int row = 1; row <= formats.Length; row++) {
            data.SetNumber(row, 1, -1234.5, formats[row - 1]);
            data.SetNumber(row, 2, 0.5, formats[row - 1]);
        }
        byte[] bytes = Save(kind, data); Assert.Equal(bytes, Save(kind, data));
        var source = IWorkSourceDocument.Open(bytes);
        IWorkTable table;
        if (kind == IWorkDocumentKind.Numbers) {
            var projection = source.ReadNumbers(); Assert.Empty(projection.SourceDeclarationIssues);
            Assert.Empty(projection.SourceReferenceIssues); Assert.Empty(projection.Diagnostics);
            table = Assert.Single(projection.Sheets[0].Tables);
        } else if (kind == IWorkDocumentKind.Pages) {
            var projection = source.ReadPages(); Assert.Empty(projection.SourceDeclarationIssues);
            Assert.Empty(projection.SourceReferenceIssues); Assert.Empty(projection.Diagnostics); table = Assert.Single(projection.Tables);
        } else {
            var projection = source.ReadKeynote(); Assert.Empty(projection.SourceDeclarationIssues);
            Assert.Empty(projection.SourceReferenceIssues); Assert.Empty(projection.Diagnostics); table = Assert.Single(projection.Slides[0].Tables);
        }
        Assert.Equal(formats.Length * 2, table.Cells.Count);
        foreach (var cell in table.Cells) {
            Assert.Equal(IWorkCellKind.Number, cell.Kind); Assert.Equal(cell.Column == 1 ? -1234.5 : 0.5, cell.Value);
            IWorkNumberFormat format = Assert.IsType<IWorkNumberFormat>(cell.NumberFormat), expected = formats[cell.Row - 1];
            Assert.Equal(expected.Kind, format.Kind); Assert.Equal(expected.DecimalPlaces, format.DecimalPlaces);
            Assert.Equal(expected.ThousandsSeparator, format.ThousandsSeparator); Assert.Equal(expected.NegativeStyle, format.NegativeStyle);
        }
    }

    [Fact]
    public void Saved_Excel_retains_numeric_values_and_explicit_format_codes() {
        var data = IWorkTableData.Create(3, 1).SetNumber(1, 1, -1234.5, IWorkNumberFormat.CreateNumber(2, true, IWorkNegativeNumberStyle.Parentheses))
            .SetNumber(2, 1, -0.125, IWorkNumberFormat.CreatePercentage(1, negativeStyle: IWorkNegativeNumberStyle.RedAndParentheses))
            .SetNumber(3, 1, 1234.5, IWorkNumberFormat.CreateScientific(4));
        using var result = IWorkSourceDocument.Open(Save(IWorkDocumentKind.Numbers, data)).ToExcelDocumentResult(
            new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED");
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using (var saved = SpreadsheetDocument.Open(output, false)) {
            Assert.Empty(new OpenXmlValidator().Validate(saved));
            var sheet = saved.WorkbookPart!.WorksheetParts.Single();
            var cells = sheet.Worksheet!.Descendants<S.Cell>().ToArray();
            var styles = saved.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            var formats = styles.NumberingFormats!.Elements<S.NumberingFormat>().ToDictionary(format => format.NumberFormatId!.Value,
                format => format.FormatCode!.Value);
            Assert.Equal(new[] { "#,##0.00;(#,##0.00)", "0.0%;[Red](0.0%)", "0.0000E+00" }, cells.Select(cell =>
                formats[styles.CellFormats!.Elements<S.CellFormat>().ElementAt((int)cell.StyleIndex!.Value).NumberFormatId!.Value]));
        }
        output.Position = 0; using var reopened = ExcelDocument.Load(output);
        Assert.Equal(-1234.5, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(-0.125, reopened.Sheets[0].CellAt(2, 1).GetValue<double>());
        Assert.Equal(1234.5, reopened.Sheets[0].CellAt(3, 1).GetValue<double>());
    }

    [Fact]
    public void Invalid_or_unsupported_formats_preserve_existing_cell_and_table_snapshot() {
        var format = IWorkNumberFormat.CreateNumber(2); var data = IWorkTableData.Create(1, 1).SetNumber(1, 1, 12.5, format);
        var document = IWorkNumbersDocument.Create(); document.AddSheet("Data").AddTable(data); byte[] baseline = document.SaveBytes();
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkNumberFormat.CreateNumber(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkNumberFormat.CreatePercentage(31));
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkNumberFormat.CreateScientific(31));
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkNumberFormat.CreateNumber(2, negativeStyle: (IWorkNegativeNumberStyle)4));
        Assert.Throws<ArgumentNullException>(() => data.SetNumber(1, 1, 99, null!));
        Assert.Throws<ArgumentOutOfRangeException>(() => data.SetNumber(1, 1, double.NaN, format));
        Assert.Throws<ArgumentOutOfRangeException>(() => data.SetNumber(2, 1, 99, format));
        var unsupported = IWorkSourceDocument.Open(Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports",
            "numbers-blank-explicit-currency-v14.5.numbers")).ReadNumbers().Sheets.SelectMany(sheet => sheet.Tables)
            .SelectMany(table => table.Cells).Select(cell => cell.NumberFormat).First(item => item?.Kind == IWorkNumberFormatKind.Currency)!;
        Assert.Throws<ArgumentException>(() => data.SetNumber(1, 1, 99, unsupported));
        Assert.Equal(12.5, data.GetCell(1, 1).Value); Assert.Same(format, data.GetCell(1, 1).NumberFormat);
        data.SetNumber(1, 1, 99, IWorkNumberFormat.CreatePercentage(4));
        Assert.Equal(baseline, document.SaveBytes());
        data.SetNumber(1, 1, 42); Assert.Null(data.GetCell(1, 1).NumberFormat);
        data.SetText(1, 1, "text"); Assert.Null(data.GetCell(1, 1).NumberFormat);
    }

    [Fact]
    public void Format_semantics_affect_identity_and_equal_independent_formats_share_deterministic_bytes() {
        byte[] Number(IWorkNumberFormat? format) {
            var data = IWorkTableData.Create(1, 1);
            if (format == null) data.SetNumber(1, 1, 1.25); else data.SetNumber(1, 1, 1.25, format);
            return Save(IWorkDocumentKind.Numbers, data);
        }
        byte[] baseline = Number(null), two = Number(IWorkNumberFormat.CreateNumber(2));
        Assert.NotEqual(baseline, two); Assert.Equal(two, Number(IWorkNumberFormat.CreateNumber(2)));
        Assert.NotEqual(two, Number(IWorkNumberFormat.CreateNumber(3)));
        Assert.NotEqual(two, Number(IWorkNumberFormat.CreateNumber(2, true)));
        Assert.NotEqual(two, Number(IWorkNumberFormat.CreateNumber(2, negativeStyle: IWorkNegativeNumberStyle.Red)));
        Assert.NotEqual(two, Number(IWorkNumberFormat.CreatePercentage(2)));
        Assert.NotEqual(two, Number(IWorkNumberFormat.CreateScientific(2)));
    }

    private static IWorkNumberFormat[] Formats() => [
        IWorkNumberFormat.CreateNumber(0), IWorkNumberFormat.CreateNumber(2, true),
        IWorkNumberFormat.CreateNumber(15), IWorkNumberFormat.CreateNumber(30),
        IWorkNumberFormat.CreateNumber(2, negativeStyle: IWorkNegativeNumberStyle.Red),
        IWorkNumberFormat.CreateNumber(2, true, IWorkNegativeNumberStyle.Parentheses),
        IWorkNumberFormat.CreateNumber(2, true, IWorkNegativeNumberStyle.RedAndParentheses),
        IWorkNumberFormat.CreatePercentage(0), IWorkNumberFormat.CreatePercentage(2, true),
        IWorkNumberFormat.CreatePercentage(30), IWorkNumberFormat.CreatePercentage(2, negativeStyle: IWorkNegativeNumberStyle.Red),
        IWorkNumberFormat.CreatePercentage(2, true, IWorkNegativeNumberStyle.Parentheses),
        IWorkNumberFormat.CreatePercentage(2, true, IWorkNegativeNumberStyle.RedAndParentheses),
        IWorkNumberFormat.CreateScientific(0), IWorkNumberFormat.CreateScientific(4), IWorkNumberFormat.CreateScientific(30)
    ];

    private static byte[] Save(IWorkDocumentKind kind, IWorkTableData data) {
        if (kind == IWorkDocumentKind.Numbers) {
            var document = IWorkNumbersDocument.Create(); document.AddSheet("Data").AddTable(data); return document.SaveBytes();
        }
        if (kind == IWorkDocumentKind.Pages) {
            var document = IWorkPagesDocument.Create(); document.AddTable(data); return document.SaveBytes();
        }
        var deck = IWorkKeynoteDocument.Create(); deck.AddSlide().AddTable(data, 0, 0, 800, 400); return deck.SaveBytes();
    }
}
