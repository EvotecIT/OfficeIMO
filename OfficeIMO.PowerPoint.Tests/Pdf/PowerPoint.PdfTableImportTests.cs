using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public class PowerPointPdfTableImportTests {
    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_ImportsDetectedTablesAsPowerPointTables() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointTableImportReport report = PowerPointPdfConverterExtensions.SaveTablesAsPowerPoint(
            LoadTables(pdf),
            presentation,
            new PdfPowerPointTableImportOptions());

        PdfPowerPointTableImportEntry result = Assert.Single(report.Entries);
        Assert.Equal(1, result.PageNumber);
        Assert.Equal(0, result.TableIndex);
        Assert.Equal(0, result.SlideIndex);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(2, result.RowCount);
        Assert.False(result.Truncated);
        Assert.True(result.HeaderRowIncluded);

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());

        A.Table table = GetSingleTable(package);
        Assert.True(table.TableProperties?.FirstRow?.Value ?? false);
        Assert.True(table.TableProperties?.BandRow?.Value ?? false);

        List<A.TableRow> rows = table.Elements<A.TableRow>().ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "Code", "Name", "Qty" }, ReadRowText(rows[0]));
        Assert.Equal(new[] { "A-100", "Alpha", "2" }, ReadRowText(rows[1]));
        Assert.Equal(new[] { "B-200", "Beta", "14" }, ReadRowText(rows[2]));
        Assert.Null(ReadHorizontalAlignment(rows[0], 2));
        Assert.Null(ReadHorizontalAlignment(rows[1], 1));
        Assert.Equal(A.TextAlignmentTypeValues.Right, ReadHorizontalAlignment(rows[1], 2));
        Assert.Equal(A.TextAlignmentTypeValues.Right, ReadHorizontalAlignment(rows[2], 2));
        long[] columnWidths = ReadColumnWidths(table);
        Assert.Equal(3, columnWidths.Length);
        Assert.True(columnWidths[1] > columnWidths[0]);
        Assert.True(columnWidths[1] > columnWidths[2]);
        Assert.Contains(ReadAllText(package), text => text == "PDF page 1, table 1");
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_AppliesRowCapsAndKeepsPresentationValidWhenEmpty() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfCore.PdfKeyValueRow.Text("Due", "2026-06-30")
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 170 },
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .PageBreak()
            .Paragraph(p => p.Text("No table on this page."))
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointTableImportReport report = PowerPointPdfConverterExtensions.SaveTablesAsPowerPoint(
            LoadTables(pdf, PdfCore.PdfPageRange.From(1, 1)),
            presentation,
            new PdfPowerPointTableImportOptions {
                MaxRows = 2,
                IncludeSourceTitles = false
            });

        PdfPowerPointTableImportEntry result = Assert.Single(report.Entries);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(3, result.TotalRowCount);
        Assert.True(result.Truncated);
        Assert.True(result.HeaderRowIncluded);
        Assert.True(report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());

        using (PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false)) {
            Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
            A.Table table = GetSingleTable(package);
            List<A.TableRow> rows = table.Elements<A.TableRow>().ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal(new[] { "Key", "Value" }, ReadRowText(rows[0]));
            Assert.Equal(new[] { "InvoiceId", "INV-001" }, ReadRowText(rows[1]));
            Assert.Equal(new[] { "Customer", "Evotec" }, ReadRowText(rows[2]));
        }

        using var emptyPresentation = new MemoryStream();
        PdfPowerPointTableImportReport emptyReport = PowerPointPdfConverterExtensions.SaveTablesAsPowerPoint(
            LoadTables(pdf, PdfCore.PdfPageRange.From(2, 2)),
            emptyPresentation,
            new PdfPowerPointTableImportOptions {
                EmptyPresentationMessage = "Nothing tabular was detected."
            });

        Assert.Empty(emptyReport.Entries);
        using PresentationDocument emptyPackage = PresentationDocument.Open(new MemoryStream(emptyPresentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(emptyPackage).ToList());
        Assert.Empty(emptyPackage.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<A.Table>()));
        Assert.Contains(ReadAllText(emptyPackage), text => text == "Nothing tabular was detected.");
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_SkipsHeaderOnlySegmentsWhenHeadersAreDisabled() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 90, 150 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointTableImportReport report = PowerPointPdfConverterExtensions.SaveTablesAsPowerPoint(
            LoadTables(pdf),
            presentation,
            new PdfPowerPointTableImportOptions {
                IncludeColumnHeaderRows = false,
                EmptyPresentationMessage = "No table rows were imported."
            });

        Assert.Empty(report.Entries);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Empty(package.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<A.Table>()));
        Assert.Contains(ReadAllText(package), text => text == "No table rows were imported.");
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_SplitsLargeTablesAcrossSlides() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 520,
                PageHeight = 420,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "C1", "C2", "C3", "C4" },
                new[] { "R1C1", "R1C2", "R1C3", "R1C4" },
                new[] { "R2C1", "R2C2", "R2C3", "R2C4" },
                new[] { "R3C1", "R3C2", "R3C3", "R3C4" },
                new[] { "R4C1", "R4C2", "R4C3", "R4C4" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 80, 80, 80, 80 },
                HeaderRowCount = 1,
                CellPaddingX = 4,
                CellPaddingY = 3
            })
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointTableImportReport report = PowerPointPdfConverterExtensions.SaveTablesAsPowerPoint(
            LoadTables(pdf),
            presentation,
            new PdfPowerPointTableImportOptions {
                MaxRowsPerSlide = 2,
                MaxColumnsPerSlide = 2
            });

        IReadOnlyList<PdfPowerPointTableImportEntry> results = report.Entries;
        Assert.Equal(4, results.Count);
        Assert.All(results, result => {
            Assert.Equal(4, result.SegmentCount);
            Assert.Equal(4, result.SourceColumnCount);
            Assert.Equal(2, result.ColumnCount);
            Assert.Equal(2, result.RowCount);
            Assert.Equal(4, result.TotalRowCount);
            Assert.True(result.HeaderRowIncluded);
        });
        Assert.Equal(new[] { 0, 1, 2, 3 }, results.Select(result => result.SegmentIndex).ToArray());
        Assert.Equal(new[] { 0, 0, 2, 2 }, results.Select(result => result.RowStartIndex).ToArray());
        Assert.Equal(new[] { 0, 2, 0, 2 }, results.Select(result => result.ColumnStartIndex).ToArray());
        Assert.Equal(new[] { 0, 1, 2, 3 }, results.Select(result => result.SlideIndex).ToArray());

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());

        List<A.Table> tables = package.PresentationPart!.SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Table>())
            .ToList();
        Assert.Equal(4, tables.Count);
        Assert.Contains(tables, table => ContainsRows(table, new[] { "C1", "C2" }, new[] { "R1C1", "R1C2" }, new[] { "R2C1", "R2C2" }));
        Assert.Contains(tables, table => ContainsRows(table, new[] { "C3", "C4" }, new[] { "R1C3", "R1C4" }, new[] { "R2C3", "R2C4" }));
        Assert.Contains(tables, table => ContainsRows(table, new[] { "C1", "C2" }, new[] { "R3C1", "R3C2" }, new[] { "R4C1", "R4C2" }));
        Assert.Contains(tables, table => ContainsRows(table, new[] { "C3", "C4" }, new[] { "R3C3", "R3C4" }, new[] { "R4C3", "R4C4" }));

        string[] text = ReadAllText(package);
        Assert.Contains(text, value => value == "PDF page 1, table 1 (part 1 of 4)");
        Assert.Contains(text, value => value == "PDF page 1, table 1 (part 4 of 4)");
    }

    private static PdfCore.PdfLogicalDocument LoadTables(byte[] pdf, params PdfCore.PdfPageRange[] ranges) {
        var layout = new PdfCore.PdfTextLayoutOptions { ForceSingleColumn = true };
        return ranges.Length == 0
            ? PdfCore.PdfLogicalDocument.Load(pdf, layout)
            : PdfCore.PdfLogicalDocument.LoadPageRanges(pdf, layout, ranges);
    }

    private static A.Table GetSingleTable(PresentationDocument package) {
        return Assert.Single(package.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<A.Table>()));
    }

    private static bool ContainsRows(A.Table table, params string[][] expectedRows) {
        string[][] rows = table.Elements<A.TableRow>()
            .Select(ReadRowText)
            .ToArray();
        if (rows.Length != expectedRows.Length) {
            return false;
        }

        for (int rowIndex = 0; rowIndex < expectedRows.Length; rowIndex++) {
            if (!rows[rowIndex].SequenceEqual(expectedRows[rowIndex])) {
                return false;
            }
        }

        return true;
    }

    private static long[] ReadColumnWidths(A.Table table) {
        return table.TableGrid!.Elements<A.GridColumn>()
            .Select(column => column.Width?.Value ?? 0L)
            .ToArray();
    }

    private static A.TextAlignmentTypeValues? ReadHorizontalAlignment(A.TableRow row, int columnIndex) {
        return row.Elements<A.TableCell>()
            .ElementAt(columnIndex)
            .TextBody?
            .Elements<A.Paragraph>()
            .FirstOrDefault()?
            .ParagraphProperties?
            .Alignment?
            .Value;
    }

    private static string[] ReadRowText(A.TableRow row) {
        return row.Elements<A.TableCell>()
            .Select(cell => string.Concat(cell.Descendants<A.Text>().Select(text => text.Text ?? string.Empty)))
            .ToArray();
    }

    private static string[] ReadAllText(PresentationDocument package) {
        return package.PresentationPart!.SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Text>())
            .Select(text => text.Text ?? string.Empty)
            .ToArray();
    }
}
