using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableDetectionValidationTests {
    [Fact]
    public void LogicalTables_DoNotTreatSparseTwoColumnParagraphsAsATable() {
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas => canvas
            .Text("Left opening paragraph", 36D, 90D, 220D, 24D, fontSize: 11D)
            .Text("Left indented paragraph", 66D, 140D, 190D, 24D, fontSize: 11D)
            .Text("Left closing paragraph", 36D, 190D, 220D, 24D, fontSize: 11D)
            .Text("Right opening paragraph", 326D, 90D, 220D, 24D, fontSize: 11D)
            .Text("Right middle paragraph", 326D, 140D, 220D, 24D, fontSize: 11D)
            .Text("Right closing paragraph", 326D, 190D, 220D, 24D, fontSize: 11D)
            .Table(new[] {
                new[] { "TABLE-MARKER", "Owner", "Amount", "Status" },
                new[] { "ACC-01", "Owner 01", "1037.25", "Approved" },
                new[] { "ACC-02", "Owner 02", "1074.50", "Review" }
            }, 36D, 300D, 510D, 150D, style: new PdfTableStyle { HeaderRowCount = 1 })))).ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        PdfLogicalTable table = Assert.Single(logical.Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("TABLE-MARKER", data.Columns.Concat(data.Rows.SelectMany(static row => row)));
    }

    [Fact]
    public void LogicalTables_RetainRegularFontProseHeavyTablesWithStableColumns() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Assigned owner", "Current workflow" },
                new[] { "North region coordinator", "Review pending requests" },
                new[] { "South region coordinator", "Approve completed requests" }
            }, style: new PdfTableStyle {
                HeaderBold = false,
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 220D, 220D }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("North region coordinator", data.Rows.SelectMany(static row => row));
        Assert.Contains("Approve completed requests", data.Rows.SelectMany(static row => row));
    }

    [Fact]
    public void LogicalTables_RetainSparseSpanningRowsWhenTheTableHasStrongEvidence() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Account", "Owner", "Amount", "Status" },
                new[] { "SECTION-A", "", "", "" },
                new[] { "ACC-01", "Owner 01", "1037.25", "Approved" }
            }, style: new PdfTableStyle { HeaderRowCount = 1 })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Contains("SECTION-A", data.Rows.SelectMany(static row => row));
        Assert.Contains("1037.25", data.Rows.SelectMany(static row => row));
    }
}
