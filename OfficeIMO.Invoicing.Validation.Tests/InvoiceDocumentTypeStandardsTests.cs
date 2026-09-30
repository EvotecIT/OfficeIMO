using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public partial class InvoiceStandardsTests {
    public static IEnumerable<object[]> DocumentTypeContracts() {
        foreach (string code in new[] { InvoiceDocumentTypes.PartialInvoice, InvoiceDocumentTypes.CorrectedInvoice, InvoiceDocumentTypes.PrepaymentInvoice, InvoiceDocumentTypes.SelfBilledInvoice }) {
            foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
                yield return new object[] { code, InvoiceSpecificationRelease.En16931_1_3_16, syntax, InvoiceProfile.En16931 };
                yield return new object[] { code, InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31, syntax, InvoiceProfile.XRechnung };
            }
            if (code != InvoiceDocumentTypes.SelfBilledInvoice)
                yield return new object[] { code, InvoiceSpecificationRelease.PeppolBis_3_0_21, InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis };
            foreach (InvoiceProfile profile in new[] { InvoiceProfile.Minimum, InvoiceProfile.BasicWithoutLines, InvoiceProfile.Basic, InvoiceProfile.En16931, InvoiceProfile.Extended })
                yield return new object[] { code, InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, profile };
        }
    }
    [InvoiceStandardsTheory]
    [MemberData(nameof(DocumentTypeContracts))]
    public async Task DocumentTypeAuthoringPassesItsSelectedAuthorityContract(string code, InvoiceSpecificationRelease release, InvoiceSyntax syntax, InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create(); invoice.TypeCode = code;
        if (code == InvoiceDocumentTypes.CorrectedInvoice) invoice.PrecedingInvoices.Add(new("ORIGINAL-2026-001", invoice.IssueDate.AddDays(-1)));
        if (profile == InvoiceProfile.PeppolBis) {
            invoice.Seller.ElectronicAddress = new("1234567890128", "0088");
            invoice.Buyer.ElectronicAddress = new("1234567890135", "0088");
        }
        var contract = new InvoiceXmlOptions(release, syntax, profile, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        byte[] xml = InvoiceSerializer.Write(invoice, contract);
        var report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, release);
        Assert.True(report.IsValid, Report(report));
        var read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal(code, read.Invoice.TypeCode);
        var rewrite = InvoiceConverter.Convert(xml, contract);
        Assert.True(rewrite.Succeeded, string.Join("; ", rewrite.Diagnostics.Select(d => d.Message)));
        Assert.Equal(code, InvoiceParser.Read(rewrite.Xml!).Invoice.TypeCode);
    }
}
