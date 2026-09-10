namespace OfficeIMO.Invoicing.Tests;

internal static class InvoiceFixture {
    internal static Invoice WithTaxCategory(string code) {
        Invoice invoice = Create();
        invoice.Lines[0].Tax = new InvoiceTaxCategory {
            Code = code, Rate = code == "O" ? null : code == "S" ? 19m : code == "L" ? 7m : 0m,
            ExemptionReason = new[] { "E", "AE", "G", "K", "O" }.Contains(code) ? "Exemption applies" : null
        };
        if (code == "O") {
            invoice.Seller.VatIdentifier = null;
            invoice.Seller.LegalRegistration = new InvoiceIdentifier("HRB 12345");
        }
        if (code == "AE" || code == "K") invoice.Buyer.VatIdentifier = "DE987654321";
        if (code == "K") invoice.Delivery = new InvoiceDelivery {
            Date = invoice.IssueDate, Address = new InvoiceAddress { CountryCode = "FR" }
        };
        return invoice;
    }

    internal static Invoice Create() {
        var invoice = new Invoice {
            Number = "INV-2026-001", IssueDate = new DateTime(2026, 9, 10), DueDate = new DateTime(2026, 10, 10),
            Currency = "EUR", BuyerReference = "04011000-12345-03", BusinessProcessId = "urn:fdc:peppol.eu:2017:poacc:billing:01:1.0",
            Seller = new InvoiceParty { Name = "Example Seller GmbH", VatIdentifier = "DE123456789", ElectronicAddress = new InvoiceIdentifier("seller@example.test", "EM"),
                Address = new InvoiceAddress { Line1 = "Seller street 1", City = "Berlin", PostCode = "10115", CountryCode = "DE" },
                Contact = new InvoiceContact { Name = "Accounts", Telephone = "+49 30 123456", Email = "seller@example.test" } },
            Buyer = new InvoiceParty { Name = "Example Buyer GmbH", ElectronicAddress = new InvoiceIdentifier("buyer@example.test", "EM"),
                Address = new InvoiceAddress { Line1 = "Buyer street 2", City = "Berlin", PostCode = "10115", CountryCode = "DE" } },
            Payment = new InvoicePayment { MeansCode = "58", Reference = "INV-2026-001" }
        };
        invoice.Payment.Accounts.Add(new InvoiceBankAccount { Identifier = "DE79000000001234567890" });
        invoice.Lines.Add(new InvoiceLine { Id = "1", Name = "Consulting", Quantity = 1m, UnitPrice = 100m, Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m } });
        return invoice;
    }
    internal static Invoice Rich(bool credit = false) {
        Invoice invoice = Create();
        invoice.TypeCode = credit ? "381" : "380";
        if (credit) invoice.DueDate = null;
        invoice.Seller.TradingName = "Example Trading";
        invoice.Seller.LegalInformation = "Registered in Berlin";
        invoice.Seller.LegalRegistration = new InvoiceIdentifier("HRB 12345");
        invoice.Seller.Identifiers.Add(new InvoiceIdentifier("seller-1"));
        invoice.Seller.TaxRegistration = "12/345/67890";
        invoice.Buyer.VatIdentifier = "DE987654321";
        invoice.Buyer.Contact = new InvoiceContact { Name = "Purchasing", Email = "purchasing@example.test" };
        invoice.Payee = new InvoiceParty { Name = "Example Payee", LegalRegistration = new InvoiceIdentifier("payee-register") };
        invoice.Payee.Identifiers.Add(new InvoiceIdentifier("payee-1"));
        invoice.TaxRepresentative = new InvoiceParty { Name = "Example Representative", VatIdentifier = "DE999999999", Address = new InvoiceAddress { CountryCode = "DE", City = "Berlin" } };
        invoice.Delivery = new InvoiceDelivery { Name = "Warehouse", Date = invoice.IssueDate, LocationIdentifier = new InvoiceIdentifier("location-1"),
            Address = new InvoiceAddress { CountryCode = "DE", Line1 = "Warehouse street 3", City = "Berlin", PostCode = "10115" } };
        invoice.Period = new InvoicePeriod { Start = invoice.IssueDate.AddDays(-10), End = invoice.IssueDate };
        invoice.TaxPointDate = invoice.IssueDate;
        invoice.ProjectReference = credit ? null : "project-1";
        invoice.ContractReference = "contract-1"; invoice.PurchaseOrderReference = "purchase-1"; invoice.SalesOrderReference = "sales-1";
        invoice.DespatchAdviceReference = "despatch-1"; invoice.ReceivingAdviceReference = "receipt-1"; invoice.TenderReference = "tender-1";
        invoice.AccountingReference = "accounting-1"; invoice.ObjectIdentifier = new InvoiceIdentifier("object-1");
        invoice.Notes.Add(new InvoiceNote("Terms of the service agreement apply", "ADU"));
        invoice.PrecedingInvoices.Add(new InvoiceReference("prior-invoice", invoice.IssueDate.AddMonths(-1)));
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument { Reference = "hours", Description = "Hours delivered", FileName = "hours.csv", MimeType = "text/csv", Data = System.Text.Encoding.UTF8.GetBytes("hours\n8\n") });
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument { Reference = "contract", ExternalUri = "https://example.test/contract" });
        invoice.PaymentTerms = "Pay within thirty days.";
        invoice.PrepaidAmount = 10m; invoice.RoundingAmount = 0.01m;
        invoice.TaxCurrency = "USD"; invoice.TaxAmountInAccountingCurrency = 22m;
        InvoiceLine line = invoice.Lines[0];
        line.Description = "Consulting delivered under contract"; line.Note = "Approved hours";
        line.Quantity = 2m; line.GrossPrice = 110m; line.PriceDiscount = 10m;
        line.Period = invoice.Period; line.OrderLineReference = "10"; line.AccountingReference = "project-cost";
        line.SellerItemIdentifier = "service-1"; line.BuyerItemIdentifier = "buyer-service-1"; line.StandardItemIdentifier = new InvoiceIdentifier("1234567890128", "0160");
        line.Classifications.Add(new InvoiceItemClassification { Value = "0721-880X", ListId = "IB" });
        line.Attributes.Add(new InvoiceItemAttribute { Name = "Service tier", Value = "Standard" });
        line.OriginCountryCode = "DE"; line.ObjectIdentifier = new InvoiceIdentifier("line-object");
        line.AllowancesAndCharges.Add(new InvoiceAllowanceCharge { Amount = 1m, Reason = "Line discount" });
        invoice.Lines.Add(new InvoiceLine { Id = "2", Name = "Books", Quantity = 2m, UnitPrice = 10m, Tax = new InvoiceTaxCategory { Code = "S", Rate = 7m } });
        invoice.AllowancesAndCharges.Add(new InvoiceAllowanceCharge { Amount = 5m, BaseAmount = 100m, Percentage = 5m, Reason = "Document discount", Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m } });
        invoice.AllowancesAndCharges.Add(new InvoiceAllowanceCharge { IsCharge = true, Amount = 2m, Reason = "Delivery", Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m } });
        return invoice;
    }
}
