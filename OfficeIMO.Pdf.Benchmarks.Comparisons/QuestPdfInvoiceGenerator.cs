using OfficeIMO.Invoicing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class QuestPdfInvoiceGenerator {
    private const string Navy = "#011534";
    private const string Accent = "#3F5CFF";
    private const string Surface = "#F5F8FC";
    private const string Border = "#D6DFEB";
    private static int _configured;

    internal static byte[] Generate(InvoiceComparisonScenario scenario) {
        Configure(scenario);
        Invoice invoice = scenario.Invoice;
        return Document.Create(document => {
            document.Page(page => {
                ConfigurePage(page);
                page.Content().Column(column => {
                    column.Spacing(12);
                    column.Item().Row(row => {
                        row.RelativeItem().Height(42).Image(scenario.LogoBytes).FitArea();
                        row.RelativeItem().AlignRight().Column(header => {
                            header.Item().AlignRight().Text("Invoice").FontSize(18).SemiBold().FontColor(Navy);
                            header.Item().AlignRight().Text(invoice.Number).FontSize(20).Bold().FontColor(Navy);
                            header.Item().AlignRight().Text($"Issued {invoice.IssueDate:dd/MM/yyyy} · Due {invoice.DueDate:dd/MM/yyyy}").FontColor("#52637A");
                        });
                    });
                    column.Item().Row(row => {
                        PartyCard(row.RelativeItem(), "SELLER", invoice.Seller);
                        row.Spacing(12);
                        PartyCard(row.RelativeItem(), "BUYER", invoice.Buyer);
                    });
                    column.Item().Table(table => {
                        table.ColumnsDefinition(columns => {
                            columns.RelativeColumn();
                            columns.RelativeColumn(1.25f);
                            columns.RelativeColumn(1.25f);
                            columns.RelativeColumn(1.9f);
                        });
                        BodyCell(table, "Currency", Surface, alignRight: false);
                        BodyCell(table, invoice.Currency, Colors.White, alignRight: false);
                        BodyCell(table, "Buyer reference", Surface, alignRight: false);
                        BodyCell(table, invoice.BuyerReference ?? string.Empty, Colors.White, alignRight: false);
                    });
                    column.Item().Table(table => {
                        table.ColumnsDefinition(columns => {
                            columns.RelativeColumn(3.7f);
                            columns.RelativeColumn(1.15f);
                            columns.RelativeColumn(1.55f);
                            columns.RelativeColumn(1f);
                            columns.RelativeColumn(1.6f);
                        });
                        Header(table, "Item", "Quantity", "Net price", "VAT", "Net amount");
                        for (int index = 0; index < invoice.Lines.Count; index++) {
                            InvoiceLine line = invoice.Lines[index];
                            string fill = index % 2 == 1 ? Surface : Colors.White;
                            BodyCell(table, $"{line.Id}. {line.Name}\n{line.Description}", fill, alignRight: false);
                            BodyCell(table, InvoiceComparisonScenario.Number(line.Quantity) + " " + line.UnitCode, fill, alignRight: true);
                            BodyCell(table, InvoiceComparisonScenario.Number(line.UnitPrice) + " / " + InvoiceComparisonScenario.Number(line.PriceBaseQuantity) + " " + line.UnitCode, fill, alignRight: true);
                            BodyCell(table, line.Tax.Code + " " + InvoiceComparisonScenario.Number(line.Tax.Rate!.Value) + "%", fill, alignRight: true);
                            BodyCell(table, InvoiceComparisonScenario.Money(scenario.Calculation.Lines[index].NetAmount, invoice.Currency), fill, alignRight: true);
                        }
                    });
                    column.Item().Table(table => {
                        table.ColumnsDefinition(columns => {
                            columns.RelativeColumn(1.2f);
                            columns.RelativeColumn(1.5f);
                            columns.RelativeColumn(1.4f);
                            columns.RelativeColumn(2f);
                        });
                        Header(table, "VAT category", "Taxable amount", "VAT amount", "Exemption");
                        BodyCell(table, "S 23%", Colors.White, alignRight: false);
                        BodyCell(table, InvoiceComparisonScenario.Money(scenario.Calculation.TaxExclusiveTotal, invoice.Currency), Colors.White, alignRight: true);
                        BodyCell(table, InvoiceComparisonScenario.Money(scenario.Calculation.TaxTotal, invoice.Currency), Colors.White, alignRight: true);
                        BodyCell(table, string.Empty, Colors.White, alignRight: false);
                    });
                });
            });
            document.Page(page => {
                ConfigurePage(page);
                page.Content().Column(column => {
                    column.Spacing(16);
                    column.Item().Row(row => {
                        row.RelativeItem(1.25f).Background(Surface).Border(0.6f).BorderColor(Border).CornerRadius(7).Padding(10).Column(payment => {
                            payment.Item().Text("Payment reference").Bold().FontColor(Navy);
                            payment.Item().Text(invoice.Number).FontColor("#52637A");
                            payment.Item().Text(invoice.PaymentTerms).FontColor("#52637A");
                        });
                        row.Spacing(18);
                        row.RelativeItem().Table(table => {
                            table.ColumnsDefinition(columns => { columns.RelativeColumn(); columns.RelativeColumn(); });
                            TotalRow(table, "Line net total", scenario.Calculation.LineNetTotal, invoice.Currency, false);
                            TotalRow(table, "Total excluding VAT", scenario.Calculation.TaxExclusiveTotal, invoice.Currency, false);
                            TotalRow(table, "VAT total", scenario.Calculation.TaxTotal, invoice.Currency, false);
                            TotalRow(table, "Total including VAT", scenario.Calculation.TaxInclusiveTotal, invoice.Currency, false);
                            TotalRow(table, "Prepaid", scenario.Calculation.PrepaidAmount, invoice.Currency, false);
                            TotalRow(table, "Amount due", scenario.Calculation.PayableAmount, invoice.Currency, true);
                        });
                    });
                    column.Item().Text("References").FontSize(15).Bold().FontColor(Navy);
                    column.Item().Table(table => {
                        table.ColumnsDefinition(columns => { columns.RelativeColumn(); columns.RelativeColumn(); });
                        BodyCell(table, "Purchase order", Colors.White, alignRight: false);
                        BodyCell(table, invoice.PurchaseOrderReference ?? string.Empty, Colors.White, alignRight: false);
                        BodyCell(table, "Document type", Colors.White, alignRight: false);
                        BodyCell(table, invoice.TypeCode, Colors.White, alignRight: false);
                    });
                    column.Item().Text("Payment").FontSize(15).Bold().FontColor(Navy);
                    column.Item().Table(table => {
                        table.ColumnsDefinition(columns => { columns.RelativeColumn(); columns.RelativeColumn(); });
                        BodyCell(table, "Payment means", Colors.White, alignRight: false);
                        BodyCell(table, $"{invoice.Payments[0].MeansCode} {invoice.Payments[0].MeansText}", Colors.White, alignRight: false);
                        BodyCell(table, "Reference", Colors.White, alignRight: false);
                        BodyCell(table, invoice.Payments[0].Reference ?? string.Empty, Colors.White, alignRight: false);
                        BodyCell(table, "IBAN", Colors.White, alignRight: false);
                        BodyCell(table, $"{invoice.Payments[0].Account!.Identifier}\n{invoice.Payments[0].Account!.Name}", Colors.White, alignRight: false);
                    });
                    column.Item().Text(invoice.PaymentTerms);
                    column.Item().Text("Notes").FontSize(15).Bold().FontColor(Navy);
                    column.Item().Text(invoice.Notes[0].Text);
                    column.Item().Text("Approvals").FontSize(15).Bold().FontColor(Navy);
                    column.Item().Row(row => {
                        Approval(row.RelativeItem(), "PREPARED BY", "Marta Nowak", "Finance · 17/09/2026");
                        row.Spacing(12);
                        Approval(row.RelativeItem(), "APPROVED BY", "Daniel Reed", "Delivery lead · 17/09/2026");
                    });
                });
            });
        }).GeneratePdf();
    }

    private static void Configure(InvoiceComparisonScenario scenario) {
        if (Interlocked.Exchange(ref _configured, 1) != 0) return;
        QuestPdfBenchmarkPolicy.ConfigureLicense();
        using var regular = new MemoryStream(scenario.RegularFont, writable: false);
        using var bold = new MemoryStream(scenario.BoldFont, writable: false);
        QuestPDF.Drawing.FontManager.RegisterFont(regular);
        QuestPDF.Drawing.FontManager.RegisterFont(bold);
    }

    private static void ConfigurePage(PageDescriptor page) {
        page.Size(QuestPDF.Helpers.PageSizes.A4);
        page.Margin(42);
        page.DefaultTextStyle(style => style.FontFamily("Carlito").FontSize(9.5f).FontColor(Navy));
    }

    private static void PartyCard(IContainer container, string label, InvoiceParty party) => container
        .Background(Surface).Border(0.6f).BorderColor(Border).CornerRadius(7).Padding(10).Column(column => {
            column.Item().Text(label).Bold().FontColor(Accent);
            column.Item().Text(party.Name).FontSize(11);
            column.Item().Text(InvoiceComparisonScenario.PartyDetails(party));
        });

    private static void Header(TableDescriptor table, params string[] values) {
        foreach (string value in values)
            table.Cell().Background(Navy).Padding(7).Text(value).Bold().FontColor(Colors.White);
    }

    private static void BodyCell(TableDescriptor table, string value, string fill, bool alignRight) {
        IContainer cell = table.Cell().Background(fill).BorderBottom(0.5f).BorderColor(Border).Padding(7);
        (alignRight ? cell.AlignRight() : cell.AlignLeft()).Text(value);
    }

    private static void TotalRow(TableDescriptor table, string label, decimal value, string currency, bool accent) {
        string fill = accent ? Accent : Colors.White;
        string text = accent ? Colors.White : Navy;
        if (accent) {
            table.Cell().Background(fill).Border(0.5f).BorderColor(Border).Padding(7).Text(label).FontColor(text).Bold();
            table.Cell().Background(fill).Border(0.5f).BorderColor(Border).Padding(7).AlignRight().Text(InvoiceComparisonScenario.Money(value, currency)).FontColor(text).Bold();
        } else {
            table.Cell().Background(fill).Border(0.5f).BorderColor(Border).Padding(7).Text(label).FontColor(text);
            table.Cell().Background(fill).Border(0.5f).BorderColor(Border).Padding(7).AlignRight().Text(InvoiceComparisonScenario.Money(value, currency)).FontColor(text);
        }
    }

    private static void Approval(IContainer container, string label, string name, string detail) => container
        .Background(Surface).Border(0.6f).BorderColor(Border).CornerRadius(7).Padding(10).Column(column => {
            column.Item().Text(label).FontColor("#52637A");
            column.Item().PaddingTop(8).BorderTop(0.5f).BorderColor(Border).Text(name).Bold();
            column.Item().Text(detail).FontColor("#52637A");
        });
}
