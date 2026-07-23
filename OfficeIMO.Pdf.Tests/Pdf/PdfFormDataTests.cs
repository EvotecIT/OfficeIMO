using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFormDataTests {
    [Fact]
    public void XfdfRoundTripPreservesScalarAndMultiValueFields() {
        byte[] source = PdfDocument.Create()
            .TextField("Customer.Name", value: "Ada")
            .MultiSelectChoiceField("Regions", new[] { "EU", "US", "APAC" }, new[] { "EU", "APAC" })
            .ToBytes();

        string xfdf = PdfDocument.Open(source).Forms.ExportXfdf();
        PdfFormDataSet parsed = PdfFormDataSet.ParseXfdf(xfdf);

        Assert.Equal(2, parsed.Fields.Count);
        Assert.Equal(new[] { "Ada" }, Assert.Single(parsed.Fields, static field => field.Name == "Customer.Name").Values);
        Assert.Equal(new[] { "EU", "APAC" }, Assert.Single(parsed.Fields, static field => field.Name == "Regions").Values);
        Assert.Contains("http://ns.adobe.com/xfdf/", xfdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ImportXfdfUsesSharedFillerAndRegeneratesReadbackValues() {
        byte[] source = PdfDocument.Create().TextField("Customer.Name", value: "Before").ToBytes();
        var data = new PdfFormDataSet(new[] { new PdfFormDataField("Customer.Name", new[] { "After" }) });

        PdfDocument updated = PdfDocument.Open(source).Forms.ImportXfdf(data.ToXfdf());
        PdfFormField field = Assert.Single(updated.Read.DocumentInfo().FormFields);

        Assert.Equal("After", field.Value);
        Assert.Single(field.Widgets);
    }

    [Fact]
    public void ParseXfdfRejectsDtdAndDuplicateNames() {
        Assert.ThrowsAny<Exception>(() => PdfFormDataSet.ParseXfdf("<!DOCTYPE xfdf [<!ENTITY x SYSTEM 'file:///tmp/x'>]><xfdf><fields><field name='A'><value>&x;</value></field></fields></xfdf>"));
        Assert.Throws<ArgumentException>(() => new PdfFormDataSet(new[] { new PdfFormDataField("A", new[] { "1" }), new PdfFormDataField("A", new[] { "2" }) }));
    }

    [Fact]
    public void ParseXfdfRejectsDocumentBeforeDomMaterializationAndImportExposesTheLimit() {
        const string xfdf = "<xfdf xmlns='http://ns.adobe.com/xfdf/'><fields><field name='A'><value>bounded</value></field></fields></xfdf>";

        Assert.Throws<InvalidOperationException>(() => PdfFormDataSet.ParseXfdf(
            xfdf,
            maxFields: 10,
            maxValueCharacters: 100,
            maxDocumentCharacters: xfdf.Length - 1));

        byte[] source = PdfDocument.Create().TextField("A", value: "before").ToBytes();
        var options = new PdfFormFillerOptions { MaxXfdfDocumentCharacters = xfdf.Length - 1 };
        Assert.Throws<InvalidOperationException>(() => PdfDocument.Open(source).Forms.ImportXfdf(xfdf, options));
    }
}
