using OfficeIMO.Invoicing;

namespace OfficeIMO.Studio.Tests;

internal static class InvoiceSample {
    internal static byte[] Create() {
        var invoice = new Invoice {
            Number = "INVOICE-2026-0042", IssueDate = new DateTime(2026, 9, 30), DueDate = new DateTime(2026, 10, 30),
            Currency = "EUR", BuyerReference = "PURCHASE-42",
            Seller = new InvoiceParty { Name = "Sample Services", Address = new InvoiceAddress { CountryCode = "DE", City = "Berlin" } },
            Buyer = new InvoiceParty { Name = "Sample Buyer", Address = new InvoiceAddress { CountryCode = "DE", City = "Berlin" } }
        };
        invoice.Seller.TaxRegistrations.Add(new InvoiceTaxRegistration("DE123456789", InvoiceTaxRegistration.VatScheme));
        invoice.Payments.Add(new InvoicePayment { MeansCode = "58", Reference = invoice.Number,
            Account = new InvoiceBankAccount { Identifier = "DE79000000001234567890" } });
        invoice.Lines.Add(new InvoiceLine { Id = "1", Name = "Consulting services", Quantity = 2, UnitPrice = 100,
            Tax = new InvoiceTaxCategory { Code = "S", Rate = 19 } });
        return InvoiceSerializer.Write(invoice, new(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Cii, InvoiceProfile.En16931));
    }
}
