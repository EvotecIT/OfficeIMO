using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceParserLimitTests {
    [Theory]
    [InlineData(InvoiceSyntax.Cii, false)]
    [InlineData(InvoiceSyntax.Cii, true)]
    [InlineData(InvoiceSyntax.Ubl, false)]
    [InlineData(InvoiceSyntax.Ubl, true)]
    public void ParsedTextMustFitTheSameModelBudgetsAsAuthoredText(InvoiceSyntax syntax, bool combined) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Notes.Add(new InvoiceNote("placeholder"));
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax))));
        XElement note = document.Descendants().Single(e => e.Name.LocalName == (syntax == InvoiceSyntax.Cii ? "IncludedNote" : "Note"));
        (syntax == InvoiceSyntax.Cii ? note.Elements().Single() : note).Value = new string('x', combined ? 900000 : 1048577);
        if (combined) for (int index = 1; index < 5; index++) note.AddAfterSelf(new XElement(note));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => InvoiceParser.Read(Encoding.UTF8.GetBytes(document.ToString())));
        Assert.Contains(combined ? "Combined invoice text" : "text value", error.Message);
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void ParsedCollectionBudgetIncludesAllGroupsAndAllowsTheBoundary(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.Create();
        for (int index = 0; index < 49997; index++) invoice.Notes.Add(new InvoiceNote("Note"));
        byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
        InvoiceReadResult boundary = InvoiceParser.Read(xml);
        Assert.True(boundary.HasCompleteMapping);
        Assert.Equal(49997, boundary.Invoice.Notes.Count);
        Assert.True(InvoiceModelValidator.Validate(boundary.Invoice).IsValid);
        Assert.Equal(xml, boundary.Write());
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(xml));
        XElement note = document.Descendants().First(e => e.Name.LocalName == (syntax == InvoiceSyntax.Cii ? "IncludedNote" : "Note"));
        note.AddAfterSelf(new XElement(note));
        Assert.Contains("50,000", Assert.Throws<InvalidDataException>(() =>
            InvoiceParser.Read(Encoding.UTF8.GetBytes(document.ToString()))).Message);
        invoice.Notes.Add(new InvoiceNote("Over budget"));
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-MODEL-LIMIT");
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void RepeatedUnmappedValuesStopAtTheDiagnosticBudget(InvoiceSyntax syntax) {
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(InvoiceFixture.Create(), new InvoiceXmlOptions(syntax))));
        for (int index = 0; index < 1001; index++) document.Root!.SetAttributeValue("extra" + index, "unmapped");
        Assert.Contains("1,000", Assert.Throws<InvalidDataException>(() =>
            InvoiceParser.Read(Encoding.UTF8.GetBytes(document.ToString()))).Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("USA")]
    [InlineData("de")]
    public void OriginCountryRequiresTwoUppercaseLetters(string country) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].OriginCountryCode = country;
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Location == "Lines[0].OriginCountryCode");
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
    }
}
