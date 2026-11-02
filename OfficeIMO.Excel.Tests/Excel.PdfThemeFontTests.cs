using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void PdfExport_ResolvesThemeFontReferencesInDrawingText() {
        using ExcelDocument document = ExcelDocument.Load(Path.Combine(_directoryDocuments, "SampleFileExcel.xlsx"));
        var options = new ExcelPdfExportOptions();

        document.ConvertToPdfWithEvidence(options).ToBytes();

        Assert.DoesNotContain(options.Report.Warnings, warning => warning.Message.Contains("+mj-lt", System.StringComparison.Ordinal));
    }
}
