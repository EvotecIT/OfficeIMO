namespace OfficeIMO.Invoicing.Tests;

internal static class Fa3InvoiceFixture {
    internal static Invoice Create() {
        var invoice = new Invoice {
            Number = "FA-2026-001", IssueDate = new DateTime(2026, 9, 30), Currency = "PLN",
            Seller = new InvoiceParty { Name = "Example Seller", Address = new InvoiceAddress { CountryCode = "PL", Line1 = "Example Street 1", Line2 = "00-001 Warszawa" } },
            Buyer = new InvoiceParty { Name = "Example Buyer", Address = new InvoiceAddress { CountryCode = "PL", Line1 = "Example Street 2", Line2 = "00-001 Warszawa" } }
        };
        invoice.Seller.TaxRegistrations.Add(new InvoiceTaxRegistration("9999999999", "NIP", InvoiceTaxRegistrationKind.Fiscal));
        invoice.Buyer.TaxRegistrations.Add(new InvoiceTaxRegistration("1111111111", "NIP", InvoiceTaxRegistrationKind.Fiscal));
        invoice.Lines.Add(new InvoiceLine { Id = "1", Name = "Consulting", Quantity = 1, UnitPrice = 100, UnitCode = "HUR", Tax = new InvoiceTaxCategory { Code = "S", Rate = 23 } });
        return invoice;
    }
    internal static Fa3InvoiceWriteOptions Options(Fa3InvoiceKind kind = Fa3InvoiceKind.TaxInvoice) =>
        new(kind, new DateTimeOffset(2026, 9, 30, 12, 0, 0, TimeSpan.Zero), new Fa3InvoiceAnnotations());
}
