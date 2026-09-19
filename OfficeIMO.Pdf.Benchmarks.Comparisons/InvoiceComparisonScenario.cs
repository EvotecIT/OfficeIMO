using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed class InvoiceComparisonScenario {
    private static readonly InvoiceXmlOptions Contract = new(
        InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
        InvoiceSyntax.Cii,
        InvoiceProfile.En16931);

    private InvoiceComparisonScenario(
        Invoice invoice,
        InvoiceCalculation calculation,
        PdfOptions pdfOptions,
        byte[] logoBytes,
        byte[] regularFont,
        byte[] boldFont) {
        Invoice = invoice;
        Calculation = calculation;
        PdfOptions = pdfOptions;
        LogoBytes = logoBytes;
        RegularFont = regularFont;
        BoldFont = boldFont;
    }

    internal Invoice Invoice { get; }
    internal InvoiceCalculation Calculation { get; }
    internal PdfOptions PdfOptions { get; }
    internal byte[] LogoBytes { get; }
    internal byte[] RegularFont { get; }
    internal byte[] BoldFont { get; }

    internal PdfInvoiceDocument CreateTypedSnapshot() => PdfInvoiceDocument.Create(Invoice, Contract, CreateLayout(LogoBytes, Invoice.IssueDate));

    internal IReadOnlyList<string> RequiredText => new[] {
        Invoice.Number,
        Invoice.Seller.Name,
        Invoice.Seller.TaxRegistrations[0].Identifier,
        Invoice.Buyer.Name,
        Invoice.IssueDate.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture),
        Invoice.DueDate!.Value.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture),
        Invoice.BuyerReference!,
        Invoice.PurchaseOrderReference!,
        Invoice.PaymentTerms!,
        Invoice.Payments[0].MeansText!,
        Invoice.Payments[0].Account!.Identifier,
        Invoice.Notes[0].Text,
        Money(Calculation.PayableAmount, Invoice.Currency),
        "Marta Nowak",
        "Daniel Reed",
        Money(Calculation.LineNetTotal, Invoice.Currency),
        Money(Calculation.TaxExclusiveTotal, Invoice.Currency),
        Money(Calculation.TaxTotal, Invoice.Currency),
        Money(Calculation.TaxInclusiveTotal, Invoice.Currency),
        Money(Calculation.PrepaidAmount, Invoice.Currency)
    }.Concat(Invoice.Lines.SelectMany((line, index) => new[] {
        line.Name,
        line.Description!,
        Number(line.Quantity) + " " + line.UnitCode,
        Number(line.UnitPrice) + " / " + Number(line.PriceBaseQuantity) + " " + line.UnitCode,
        line.Tax.Code + " " + Number(line.Tax.Rate!.Value) + "%",
        Money(Calculation.Lines[index].NetAmount, Invoice.Currency)
    })).ToArray();

    internal static InvoiceComparisonScenario Create() {
        Invoice invoice = CreateInvoice();
        InvoiceCalculation calculation = InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        byte[] logo = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Brand", "evotec-logo-horizontal-gradient-2400.png"));
        byte[] regular = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "Carlito-Regular.ttf"));
        byte[] bold = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "Carlito-Bold.ttf"));
        var options = new PdfOptions { DefaultFontSize = 9.5D }
            .UseFontFamily(new PdfEmbeddedFontFamily("Carlito", regular, bold));
        return new InvoiceComparisonScenario(invoice, calculation, options, logo, regular, bold);
    }

    internal static string PartyDetails(InvoiceParty party) => string.Join("\n", new[] {
        party.Address.Line1,
        string.Join(" ", new[] { party.Address.PostCode, party.Address.City }.Where(value => !string.IsNullOrWhiteSpace(value))),
        party.Address.CountryCode
    }.Concat(party.TaxRegistrations.Select(registration => registration.SchemeId + ": " + registration.Identifier))
        .Where(value => !string.IsNullOrWhiteSpace(value)));

    internal static string Number(decimal value) => value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);

    internal static string Money(decimal value, string currency) =>
        value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) + " " + currency;

    private static InvoicePdfLayoutOptions CreateLayout(byte[] logo, DateTime? issueDate = null) {
        var layout = InvoicePdfLayoutOptions.ForCultures("en-GB");
        layout.Theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));
        layout.LogoBytes = logo;
        layout.LogoAlternativeText = "Evotec";
        DateTime approvalDate = issueDate ?? new DateTime(2026, 9, 17);
        layout.Approvals.Add(new InvoicePdfApproval("Prepared by", "Marta Nowak", "Finance", approvalDate));
        layout.Approvals.Add(new InvoicePdfApproval("Approved by", "Daniel Reed", "Delivery lead", approvalDate));
        return layout;
    }

    private static Invoice CreateInvoice() {
        var invoice = new Invoice {
            Number = "EVO-DEMO-2026-0917",
            IssueDate = new DateTime(2026, 9, 17),
            DueDate = new DateTime(2026, 10, 1),
            Currency = "EUR",
            BuyerReference = "NW-DOC-2048",
            PurchaseOrderReference = "PO-2048-09",
            PaymentTerms = "Payment due within 14 days. Use the invoice number as the transfer reference.",
            Seller = new InvoiceParty {
                Name = "Evotec Services sp. z o.o.",
                Address = new InvoiceAddress { Line1 = "Demonstration address", City = "Warsaw", PostCode = "00-001", CountryCode = "PL" }
            },
            Buyer = new InvoiceParty {
                Name = "Northwind Field Operations",
                Address = new InvoiceAddress { Line1 = "12 Harbour Street", City = "Copenhagen", PostCode = "1050", CountryCode = "DK" }
            }
        };
        invoice.Seller.TaxRegistrations.Add(new InvoiceTaxRegistration("PL0000000000", InvoiceTaxRegistration.VatScheme));
        invoice.Payments.Add(new InvoicePayment {
            MeansCode = "58",
            MeansText = "SEPA credit transfer",
            Reference = invoice.Number,
            Account = new InvoiceBankAccount { Identifier = "PL10105000997603123456789123", Name = "Evotec Services" }
        });
        AddLine(invoice, "1", "Document workflow discovery", "Architecture workshop and conversion inventory", 6m, 145m);
        AddLine(invoice, "2", "Invoice automation implementation", "Typed model, PDF layout and embedded CII delivery", 18m, 145m);
        AddLine(invoice, "3", "Validation evidence", "Artifact inspection, semantic readback and release notes", 7.5m, 145m);
        AddLine(invoice, "4", "Team enablement", "Developer handover and API walkthrough", 4m, 145m);
        invoice.Notes.Add(new InvoiceNote("Demonstration document with non-operational seller and bank details."));
        return invoice;
    }

    private static void AddLine(Invoice invoice, string id, string name, string description, decimal quantity, decimal unitPrice) =>
        invoice.Lines.Add(new InvoiceLine {
            Id = id,
            Name = name,
            Description = description,
            Quantity = quantity,
            UnitCode = "HUR",
            UnitPrice = unitPrice,
            Tax = new InvoiceTaxCategory { Code = "S", Rate = 23m }
        });
}
