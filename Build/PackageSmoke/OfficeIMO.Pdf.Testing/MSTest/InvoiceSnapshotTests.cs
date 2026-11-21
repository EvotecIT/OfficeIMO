using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeIMO.Pdf.Testing;

[TestClass]
public class InvoiceSnapshotTests {
    [TestMethod]
    public void Invoice() => PdfSnapshot.Verify(InvoiceSample.Create(), "invoice");
}
