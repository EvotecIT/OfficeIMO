using OfficeIMO.Pdf.Testing;
using Xunit;

public class InvoiceSnapshotTests {
    [Fact]
    public void Invoice() => PdfSnapshot.Verify(InvoiceSample.Create(), "invoice");
}
