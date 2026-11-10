using NUnit.Framework;
using OfficeIMO.Pdf.Testing;

public class InvoiceSnapshotTests {
    [Test]
    public void Invoice() => PdfSnapshot.Verify(InvoiceSample.Create(), "invoice");
}
