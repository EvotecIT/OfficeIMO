using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Benchmarks;

internal sealed class InvoiceBenchmarkCorpus {
    internal const int LineCount = 25;
    internal const string Marker = "FINAL-BENCHMARK-MARKER";
    internal static readonly InvoiceXmlOptions Contract = new(
        InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
        InvoiceSyntax.Cii,
        InvoiceProfile.En16931);

    private InvoiceBenchmarkCorpus(Invoice invoice, byte[] xml, PdfOptions pdfOptions) {
        Invoice = invoice;
        Xml = xml;
        PdfOptions = pdfOptions;
    }

    internal Invoice Invoice { get; }
    internal byte[] Xml { get; }
    internal PdfOptions PdfOptions { get; }

    internal static InvoiceBenchmarkCorpus Create() {
        var invoice = new Invoice {
            Number = "BENCH-2026-001",
            IssueDate = new DateTime(2026, 9, 15),
            DueDate = new DateTime(2026, 10, 15),
            Currency = "EUR",
            BuyerReference = "BENCHMARK-BUYER",
            Seller = new InvoiceParty {
                Name = "Benchmark Seller GmbH",
                Address = new InvoiceAddress { Line1 = "Teststraße 1", City = "Berlin", PostCode = "10115", CountryCode = "DE" }
            },
            Buyer = new InvoiceParty {
                Name = "Nabywca 日本語 العربية हिन्दी",
                Address = new InvoiceAddress { Line1 = "Ulica testowa 2", City = "Warszawa", PostCode = "00-001", CountryCode = "PL" }
            },
            PaymentTerms = "Thirty days / trzydzieści dni"
        };
        invoice.Seller.TaxRegistrations.Add(new InvoiceTaxRegistration("DE123456789", InvoiceTaxRegistration.VatScheme));
        invoice.Payments.Add(new InvoicePayment {
            MeansCode = "58",
            Reference = invoice.Number,
            Account = new InvoiceBankAccount { Identifier = "DE79000000001234567890" }
        });
        for (int index = 1; index <= LineCount; index++) {
            invoice.Lines.Add(new InvoiceLine {
                Id = index.ToString(System.Globalization.CultureInfo.InvariantCulture),
                Name = "Consulting / Doradztwo 日本語 العربية हिन्दी " + index,
                Description = string.Join(" ", Enumerable.Repeat("Long multilingual service description", 8)) +
                    (index == LineCount ? " " + Marker : string.Empty),
                Quantity = index,
                UnitPrice = 12.345m,
                Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m }
            });
        }
        invoice.Notes.Add(new InvoiceNote("Benchmark corpus uses one deterministic semantic invoice across every operation."));
        InvoiceModelValidator.Validate(invoice).ThrowIfInvalid();
        byte[] xml = InvoiceSerializer.Write(invoice, Contract);
        return new InvoiceBenchmarkCorpus(invoice, xml, CreatePdfOptions());
    }

    internal PdfInvoiceDocument CapturePdf() => PdfInvoiceDocument.Create(
        Invoice,
        Contract,
        InvoicePdfLayoutOptions.ForCultures("pl-PL", "en-GB"));

    private static PdfOptions CreatePdfOptions() {
        string fonts = Path.Combine(AppContext.BaseDirectory, "Fonts");
        byte[] baseFont = File.ReadAllBytes(Path.Combine(fonts, "SourceSerif4-Regular.otf"));
        return new PdfOptions { IncludeStandardFontToUnicodeMaps = true }
            .EmbedStandardFont(PdfStandardFont.Helvetica, baseFont, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, baseFont, "OfficeIMO Source Serif")
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Noto Sans Japanese", File.ReadAllBytes(Path.Combine(fonts, "NotoSansJP-OfficeIMO-Common.ttf"))),
                new PdfEmbeddedFontFallbackCandidate("Noto Sans Arabic", File.ReadAllBytes(Path.Combine(fonts, "NotoSansArabic-Regular.ttf"))),
                new PdfEmbeddedFontFallbackCandidate("Noto Sans Devanagari", File.ReadAllBytes(Path.Combine(fonts, "NotoSansDevanagari-Regular.ttf")))
            }));
    }
}
