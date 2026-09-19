using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a polished PDF invoice and embeds its CII representation from the typed invoice model.</summary>
internal static class BrandedInvoice {
    internal static void Create(string folder) {
        Invoice invoice = CreateInvoice();
        var contract = new InvoiceXmlOptions(
            InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii,
            InvoiceProfile.En16931);
        var layout = InvoicePdfLayoutOptions.ForCultures("en-GB");
        layout.Theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));
        layout.LogoBytes = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory,
            "Assets",
            "Brand",
            "Evotec",
            "evotec-logo-horizontal-gradient-2400.png"));
        layout.LogoAlternativeText = "Evotec";
        layout.Approvals.Add(new InvoicePdfApproval("Prepared by", "Marta Nowak", "Finance", invoice.IssueDate));
        layout.Approvals.Add(new InvoicePdfApproval("Approved by", "Daniel Reed", "Delivery lead", invoice.IssueDate));

        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, contract, layout);
        string fonts = Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts");
        var pdfOptions = new PdfOptions {
            DefaultFontSize = 9.5D,
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
        }.UseFontFamily(new PdfEmbeddedFontFamily(
            "Carlito",
            File.ReadAllBytes(Path.Combine(fonts, "Carlito-Regular.ttf")),
            File.ReadAllBytes(Path.Combine(fonts, "Carlito-Bold.ttf"))));

        File.WriteAllBytes(Path.Combine(folder, "example.xml"), snapshot.ToXmlBytes());
        File.WriteAllBytes(Path.Combine(folder, "example.pdf"), snapshot.ToPdfBytes(pdfOptions));
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
                ElectronicAddress = new InvoiceIdentifier("contact@evotec.pl", "EM"),
                Address = new InvoiceAddress {
                    Line1 = "Demonstration address",
                    City = "Warsaw",
                    PostCode = "00-001",
                    CountryCode = "PL"
                },
                Contact = new InvoiceContact { Name = "Finance team", Email = "contact@evotec.pl" }
            },
            Buyer = new InvoiceParty {
                Name = "Northwind Field Operations",
                ElectronicAddress = new InvoiceIdentifier("accounts@example.test", "EM"),
                Address = new InvoiceAddress {
                    Line1 = "12 Harbour Street",
                    City = "Copenhagen",
                    PostCode = "1050",
                    CountryCode = "DK"
                },
                Contact = new InvoiceContact { Name = "Accounts payable", Email = "accounts@example.test" }
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
        invoice.Notes.Add(new InvoiceNote("Demonstration document. Seller identifiers, address and bank account are intentionally non-operational sample data."));
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
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
