using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInspectionProjectionScopeTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UninspectedVisualScopeCannotReportKnownContentAbsence(bool documentOverload) {
        PdfDocument pdf = PdfDocument.Load(SingleStream("1 0 0 rg 40 40 20 20 re f"));
        PdfDocumentReadResult full = pdf.Read();
        Assert.Equal(1, Assert.Single(full.Pages).VectorPrimitiveCount);
        Assert.True(PdfLogicalTableAnalysis.AnalyzeExtractionScope(full).HasOmittedPageContent);
        PdfDocumentReadResult lean = pdf.Read(new PdfLoadOptions { IncludeImages = false });
        PdfLogicalPage page = Assert.Single(lean.Pages);
        Assert.False(page.HasVisualPrimitiveCounts);
        Assert.False(page.IncludesImages);
        Assert.True(page.IncludesFormWidgets);
        Assert.Throws<InvalidOperationException>(() => documentOverload
            ? PdfLogicalTableAnalysis.AnalyzeExtractionScope(lean)
            : PdfLogicalTableAnalysis.AnalyzeExtractionScope(lean.Pages));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnrequestedWidgetProjectionCannotEstablishOmissionScope(bool documentOverload) {
        PdfDocumentReadResult lean = PdfDocument.Load(SingleStream(string.Empty)).Read(new PdfLoadOptions { IncludeFormWidgets = false });
        PdfLogicalPage page = Assert.Single(lean.Pages);
        Assert.True(page.HasVisualPrimitiveCounts);
        Assert.True(page.IncludesImages);
        Assert.False(page.IncludesFormWidgets);
        Assert.Throws<InvalidOperationException>(() => documentOverload
            ? PdfLogicalTableAnalysis.AnalyzeExtractionScope(lean)
            : PdfLogicalTableAnalysis.AnalyzeExtractionScope(lean.Pages));
    }

    [Fact]
    public void CompleteEmptyProjectionCanStillEstablishKnownAbsence() {
        PdfDocumentReadResult full = PdfDocument.Load(SingleStream(string.Empty)).Read();
        Assert.True(Assert.Single(full.Pages).IncludesImages);
        Assert.True(Assert.Single(full.Pages).IncludesFormWidgets);
        Assert.False(PdfLogicalTableAnalysis.AnalyzeExtractionScope(full).HasOmittedPageContent);
        Assert.False(PdfLogicalTableAnalysis.AnalyzeExtractionScope(full.Pages).HasOmittedPageContent);
    }

    [Fact]
    public void SelectedRepeatedPagesPreserveUninspectedProjectionAndPermitTableExtraction() {
        PdfDocumentReadResult lean = PdfDocument.Load(SingleStream("1 0 0 rg 40 40 20 20 re f"))
            .Read(new PdfLoadOptions { IncludeImages = false, IncludeFormWidgets = false,
                PageSelection = PdfPageSelection.From(1, 1) });
        Assert.Equal(2, lean.Pages.Count);
        Assert.All(lean.Pages, page => {
            Assert.False(page.IncludesImages);
            Assert.False(page.IncludesFormWidgets);
            Assert.False(page.HasVisualPrimitiveCounts);
        });
        Assert.Throws<InvalidOperationException>(() => PdfLogicalTableAnalysis.AnalyzeExtractionScope(lean));
        Assert.Throws<InvalidOperationException>(() => PdfLogicalTableAnalysis.AnalyzeExtractionScope(lean.Pages));
        Assert.Empty(PdfLogicalTableAnalysis.ExtractTables(lean));
    }

    private static byte[] SingleStream(string content) {
        string[] objects = { "<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 200] /Resources << >> /Contents 4 0 R >>",
            "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream" };
        var bytes = new StringBuilder("%PDF-1.7\n"); var offsets = new List<int>();
        for (int i = 0; i < objects.Length; i++) { offsets.Add(bytes.Length); bytes.Append(i + 1).Append(" 0 obj\n").Append(objects[i]).Append("\nendobj\n"); }
        int xref = bytes.Length; bytes.Append("xref\n0 5\n0000000000 65535 f \n");
        foreach (int offset in offsets) bytes.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        bytes.Append("trailer\n<< /Root 1 0 R /Size 5 >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(bytes.ToString());
    }
}
