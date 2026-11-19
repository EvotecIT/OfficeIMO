#if NET10_0_OR_GREATER
using OfficeIMO.Invoicing;
using OnePdf.Layout;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Lays out the shared two-page invoice contract with OnePDF.Layout, following the same structure as the QuestPDF
/// lane: header with logo, party cards, references, line items, VAT summary, totals, payment, notes and approvals.
/// OnePDF.Layout writes a tagged PDF by default.
/// </summary>
internal static class OnePdfInvoiceGenerator {
    private const double FontSize = 9.5;
    private static readonly LayoutColor Navy = LayoutColor.FromHex("#011534");
    private static readonly LayoutColor Accent = LayoutColor.FromHex("#3F5CFF");
    private static readonly LayoutColor Surface = LayoutColor.FromHex("#F5F8FC");
    private static readonly LayoutColor Border = LayoutColor.FromHex("#D6DFEB");
    private static readonly LayoutColor Muted = LayoutColor.FromHex("#52637A");
    private static readonly LayoutColor White = LayoutColor.White;

    internal static byte[] Generate(InvoiceComparisonScenario scenario) {
        Invoice invoice = scenario.Invoice;
        return LayoutDocument.Compose(document => {
            document.Title("Invoice " + invoice.Number);
            document.Fonts.Add(scenario.RegularFont, "Carlito");
            document.Fonts.Add(scenario.BoldFont, "Carlito");
            document.Fonts.DefaultFamily = "Carlito";

            document.Page(page => {
                ConfigurePage(page);
                page.Content().Column(column => {
                    column.Spacing(12);
                    column.Item().Row(row => {
                        row.RelativeItem().Height(42).Image(scenario.LogoBytes).Contain().AltText("Evotec");
                        row.RelativeItem().AlignRight().Column(header => {
                            Text(header.Item().AlignRight(), "Invoice").FontSize(18).Weight(600);
                            Text(header.Item().AlignRight(), invoice.Number).FontSize(20).Bold();
                            Text(header.Item().AlignRight(), $"Issued {invoice.IssueDate:dd/MM/yyyy} · Due {invoice.DueDate:dd/MM/yyyy}").FontColor(Muted);
                        });
                    });
                    column.Item().Row(row => {
                        row.Spacing(12);
                        PartyCard(row.RelativeItem(), "SELLER", invoice.Seller);
                        PartyCard(row.RelativeItem(), "BUYER", invoice.Buyer);
                    });
                    column.Item().Table(table => {
                        table.Columns(columns => {
                            columns.Relative();
                            columns.Relative(1.25);
                            columns.Relative(1.25);
                            columns.Relative(1.9);
                        });
                        table.Row(row => {
                            BodyCell(row, "Currency", Surface, alignRight: false);
                            BodyCell(row, invoice.Currency, White, alignRight: false);
                            BodyCell(row, "Buyer reference", Surface, alignRight: false);
                            BodyCell(row, invoice.BuyerReference ?? string.Empty, White, alignRight: false);
                        });
                    });
                    column.Item().Table(table => {
                        table.Columns(columns => {
                            columns.Relative(3.7);
                            columns.Relative(1.15);
                            columns.Relative(1.55);
                            columns.Relative(1);
                            columns.Relative(1.6);
                        });
                        Header(table, "Item", "Quantity", "Net price", "VAT", "Net amount");
                        for (int index = 0; index < invoice.Lines.Count; index++) {
                            InvoiceLine line = invoice.Lines[index];
                            LayoutColor fill = index % 2 == 1 ? Surface : White;
                            decimal netAmount = scenario.Calculation.Lines[index].NetAmount;
                            table.Row(row => {
                                BodyCell(row, $"{line.Id}. {line.Name}\n{line.Description}", fill, alignRight: false);
                                BodyCell(row, InvoiceComparisonScenario.Number(line.Quantity) + " " + line.UnitCode, fill, alignRight: true);
                                BodyCell(row, InvoiceComparisonScenario.Number(line.UnitPrice) + " / " + InvoiceComparisonScenario.Number(line.PriceBaseQuantity) + " " + line.UnitCode, fill, alignRight: true);
                                BodyCell(row, line.Tax.Code + " " + InvoiceComparisonScenario.Number(line.Tax.Rate!.Value) + "%", fill, alignRight: true);
                                BodyCell(row, InvoiceComparisonScenario.Money(netAmount, invoice.Currency), fill, alignRight: true);
                            });
                        }
                    });
                    column.Item().Table(table => {
                        table.Columns(columns => {
                            columns.Relative(1.2);
                            columns.Relative(1.5);
                            columns.Relative(1.4);
                            columns.Relative(2);
                        });
                        Header(table, "VAT category", "Taxable amount", "VAT amount", "Exemption");
                        table.Row(row => {
                            BodyCell(row, "S 23%", White, alignRight: false);
                            BodyCell(row, InvoiceComparisonScenario.Money(scenario.Calculation.TaxExclusiveTotal, invoice.Currency), White, alignRight: true);
                            BodyCell(row, InvoiceComparisonScenario.Money(scenario.Calculation.TaxTotal, invoice.Currency), White, alignRight: true);
                            BodyCell(row, string.Empty, White, alignRight: false);
                        });
                    });
                });
            });

            document.Page(page => {
                ConfigurePage(page);
                page.Content().Column(column => {
                    column.Spacing(16);
                    column.Item().Row(row => {
                        row.Spacing(18);
                        row.RelativeItem(1.25).Background(Surface).Border(0.6, Border).Rounded(7).Padding(10).Column(payment => {
                            Text(payment.Item(), "Payment reference").Bold();
                            Text(payment.Item(), invoice.Number).FontColor(Muted);
                            Text(payment.Item(), invoice.PaymentTerms ?? string.Empty).FontColor(Muted);
                        });
                        row.RelativeItem().Table(table => {
                            table.Columns(columns => {
                                columns.Relative();
                                columns.Relative();
                            });
                            TotalRow(table, "Line net total", scenario.Calculation.LineNetTotal, invoice.Currency, false);
                            TotalRow(table, "Total excluding VAT", scenario.Calculation.TaxExclusiveTotal, invoice.Currency, false);
                            TotalRow(table, "VAT total", scenario.Calculation.TaxTotal, invoice.Currency, false);
                            TotalRow(table, "Total including VAT", scenario.Calculation.TaxInclusiveTotal, invoice.Currency, false);
                            TotalRow(table, "Prepaid", scenario.Calculation.PrepaidAmount, invoice.Currency, false);
                            TotalRow(table, "Amount due", scenario.Calculation.PayableAmount, invoice.Currency, true);
                        });
                    });
                    SectionHeading(column.Item(), "References");
                    TwoColumnTable(column.Item(),
                        ("Purchase order", invoice.PurchaseOrderReference ?? string.Empty),
                        ("Document type", invoice.TypeCode));
                    SectionHeading(column.Item(), "Payment");
                    InvoicePayment payment = invoice.Payments[0];
                    TwoColumnTable(column.Item(),
                        ("Payment means", $"{payment.MeansCode} {payment.MeansText}"),
                        ("Reference", payment.Reference ?? string.Empty),
                        ("IBAN", $"{payment.Account!.Identifier}\n{payment.Account!.Name}"));
                    Text(column.Item(), invoice.PaymentTerms ?? string.Empty);
                    SectionHeading(column.Item(), "Notes");
                    Text(column.Item(), invoice.Notes[0].Text);
                    SectionHeading(column.Item(), "Approvals");
                    column.Item().Row(row => {
                        row.Spacing(12);
                        Approval(row.RelativeItem(), "PREPARED BY", "Marta Nowak", "Finance · 17/09/2026");
                        Approval(row.RelativeItem(), "APPROVED BY", "Daniel Reed", "Delivery lead · 17/09/2026");
                    });
                });
            });
        }).ToBytes();
    }

    private static void ConfigurePage(PageComposer page) {
        page.Size(PageSize.A4);
        page.Margin(42);
    }

    private static TextDescriptor Text(Slot slot, string value) =>
        slot.Text(value).FontSize(FontSize).FontColor(Navy);

    private static void SectionHeading(Slot slot, string value) => Text(slot, value).FontSize(15).Bold();

    private static void PartyCard(Slot slot, string label, InvoiceParty party) =>
        slot.Background(Surface).Border(0.6, Border).Rounded(7).Padding(10).Column(column => {
            Text(column.Item(), label).Bold().FontColor(Accent);
            Text(column.Item(), party.Name).FontSize(11);
            Text(column.Item(), InvoiceComparisonScenario.PartyDetails(party));
        });

    private static void Header(TableDescriptor table, params string[] values) =>
        table.HeaderRow(row => {
            foreach (string value in values) {
                Text(row.Cell().CellBackground(Navy).CellPadding(7, 7, 7, 7), value).Bold().FontColor(White);
            }
        });

    private static void BodyCell(TableRowDescriptor row, string value, LayoutColor fill, bool alignRight) {
        TableCellSlot cell = row.Cell().CellBackground(fill).CellBorderBottom(0.5, Border).CellPadding(7, 7, 7, 7);
        Text(alignRight ? cell.AlignRight() : cell.AlignLeft(), value);
    }

    private static void TwoColumnTable(Slot slot, params (string Label, string Value)[] rows) =>
        slot.Table(table => {
            table.Columns(columns => {
                columns.Relative();
                columns.Relative();
            });
            foreach ((string label, string value) in rows) {
                table.Row(row => {
                    BodyCell(row, label, White, alignRight: false);
                    BodyCell(row, value, White, alignRight: false);
                });
            }
        });

    private static void TotalRow(TableDescriptor table, string label, decimal value, string currency, bool accent) {
        LayoutColor fill = accent ? Accent : White;
        LayoutColor text = accent ? White : Navy;
        table.Row(row => {
            TextDescriptor labelText = Text(row.Cell().CellBackground(fill).CellBorder(0.5, Border).CellPadding(7, 7, 7, 7), label).FontColor(text);
            TextDescriptor amountText = Text(row.Cell().CellBackground(fill).CellBorder(0.5, Border).CellPadding(7, 7, 7, 7).AlignRight(), InvoiceComparisonScenario.Money(value, currency)).FontColor(text);
            if (accent) {
                labelText.Bold();
                amountText.Bold();
            }
        });
    }

    private static void Approval(Slot slot, string label, string name, string detail) =>
        slot.Background(Surface).Border(0.6, Border).Rounded(7).Padding(10).Column(column => {
            Text(column.Item(), label).FontColor(Muted);
            Text(column.Item().PaddingTop(8).BorderTop(0.5, Border), name).Bold();
            Text(column.Item(), detail).FontColor(Muted);
        });
}
#endif
