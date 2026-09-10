using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceParsingTests {
    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void WrittenInvoicesCanBeEditedAndConvertedWithoutLosingBusinessFields(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceCalculationTests.Example();
        invoice.Notes.Add(new InvoiceNote("Contract terms apply", "ADU"));
        invoice.Period = new InvoicePeriod { Start = invoice.IssueDate, End = invoice.IssueDate.AddDays(10) };
        invoice.Lines[0].Classifications.Add(new InvoiceItemClassification { Value = "123", ListId = "ZZZ", ListVersion = "1" });
        invoice.Lines[0].Attributes.Add(new InvoiceItemAttribute { Name = "Colour", Value = "Blue" });
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument { Reference = "terms", Description = "Terms", Data = new byte[] { 1, 2, 3 }, FileName = "terms.bin", MimeType = "application/octet-stream" });
        byte[] original = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
        InvoiceReadResult read = InvoiceParser.Read(original);
        Assert.True(read.HasCompleteMapping, string.Join("\n", read.UnmappedData.Select(d => d.Location + ": " + d.Message)));
        Assert.Equal(invoice.Number, read.Invoice.Number);
        Assert.Equal(invoice.Lines[0].Attributes[0].Value, read.Invoice.Lines[0].Attributes[0].Value);
        Assert.Equal(invoice.SupportingDocuments[0].Data, read.Invoice.SupportingDocuments[0].Data);
        Assert.Equal(original, read.GetOriginalBytes());
        read.Invoice.Number = "EDITED";
        Assert.Equal("EDITED", InvoiceParser.Read(read.Write()).Invoice.Number);
        Assert.Equal(original, read.GetOriginalBytes());
        var result = InvoiceConverter.Convert(original, new InvoiceXmlOptions(syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii));
        Assert.True(result.Succeeded, string.Join("\n", result.Diagnostics.Select(d => d.Location + ": " + d.Message)));
        Assert.Equal(read.Invoice.DeclaredTotals!.PayableAmount, InvoiceParser.Read(result.Xml!).Invoice.DeclaredTotals!.PayableAmount);
    }

    [Theory]
    [InlineData("01.01a-INVOICE_ubl.xml", "123456XX", "336.9")]
    [InlineData("01.01a-INVOICE_uncefact.xml", "123456XX", "336.9")]
    public void IndependentInvoicesPreserveDeclaredAmountsAndMapTheirBusinessData(string name, string number, string due) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "KoSIT", name));
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.Equal(number, read.Invoice.Number);
        Assert.Equal(decimal.Parse(due, System.Globalization.CultureInfo.InvariantCulture), read.Invoice.DeclaredTotals!.PayableAmount);
        Assert.True(read.HasCompleteMapping, string.Join("\n", read.UnmappedData.Select(d => d.Location + ": " + d.Message)));
        InvoiceModelValidationResult validation = InvoiceModelValidator.Validate(read.Invoice);
        Assert.True(validation.IsValid, string.Join("\n", validation.Diagnostics.Select(d => d.Location + ": " + d.Message)));
        var target = new InvoiceXmlOptions(read.Declaration.Syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii, InvoiceProfile.XRechnung);
        InvoiceConversionResult result = InvoiceConverter.Convert(source, target);
        Assert.True(result.Succeeded, string.Join("\n", result.Diagnostics.Select(d => d.Location + ": " + d.Message)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnknownElementsAndAttributesBlockRewritingAndConversion(bool attribute) {
        byte[] original = InvoiceSerializer.Write(InvoiceCalculationTests.Example());
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(original));
        XNamespace extension = "urn:unmapped:test";
        if (attribute) document.Root!.SetAttributeValue(extension + "businessFlag", "important");
        else document.Root!.Add(new XElement(extension + "BusinessData", "important"));
        byte[] source = Encoding.UTF8.GetBytes(document.ToString());
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.False(read.HasCompleteMapping);
        Assert.Contains(read.UnmappedData, d => d.Code == "INV-UNMAPPED");
        Assert.Equal(source, read.GetOriginalBytes());
        Assert.Throws<InvalidDataException>(() => read.Write());
        var conversion = InvoiceConverter.Convert(source, new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        Assert.False(conversion.Succeeded);
        Assert.Null(conversion.Xml);
    }

    [Fact]
    public void DuplicateAmountsAndInvalidNumbersAreNotSilentlySelectedOrRounded() {
        string source = Encoding.UTF8.GetString(InvoiceSerializer.Write(InvoiceCalculationTests.Example()));
        Assert.Throws<InvalidDataException>(() => InvoiceParser.Read(Encoding.UTF8.GetBytes(source.Replace("<ram:DuePayableAmount>119.00</ram:DuePayableAmount>",
            "<ram:DuePayableAmount>119.00</ram:DuePayableAmount><ram:DuePayableAmount>0</ram:DuePayableAmount>"))));
        Assert.Throws<InvalidDataException>(() => InvoiceParser.Read(Encoding.UTF8.GetBytes(source.Replace("<ram:ChargeAmount>100</ram:ChargeAmount>", "<ram:ChargeAmount>1e1000</ram:ChargeAmount>"))));
    }
}
