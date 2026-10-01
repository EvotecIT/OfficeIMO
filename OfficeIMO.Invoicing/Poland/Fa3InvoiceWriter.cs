using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

/// <summary>Native FA(3) authoring over the shared invoice and explicit national declarations.</summary>
/// <remarks>Writing does not establish fiscal correctness or KSeF acceptance. Validate the exact result against the pinned national schema.</remarks>
public static partial class Fa3InvoiceWriter {
    private static readonly XNamespace Ns = Fa3InvoiceReader.NamespaceUri;
    private static readonly string[] TaxSuffixes = { "1", "2", "3", "4", "5", "6_1", "6_2", "6_3", "7", "8", "9", "10", "11" };

    /// <summary>Inspects the bounded supported mapping and arithmetic without writing or discarding populated fields.</summary>
    public static IReadOnlyList<InvoiceDiagnostic> Inspect(Invoice invoice, Fa3InvoiceWriteOptions options) => Prepare(invoice, options).Diagnostics;

    /// <summary>Writes UTF-8 FA(3), schema 1-0E, variant 3. Creation time and national declarations are supplied explicitly.</summary>
    public static byte[] Write(Invoice invoice, Fa3InvoiceWriteOptions options) {
        Prepared prepared = Prepare(invoice, options);
        if (prepared.Diagnostics.Any(item => item.Severity == InvoiceDiagnosticSeverity.Error))
            throw new InvalidDataException(string.Join(Environment.NewLine, prepared.Diagnostics.Select(item => item.Location + ": " + item.Message)));
        Fa3FiscalAmounts amounts = prepared.Amounts!;
        var fa = new XElement(Ns + "Fa",
            Element("KodWaluty", invoice.Currency), Element("P_1", Date(invoice.IssueDate)),
            Element("P_1M", options.PlaceOfIssue), Element("P_2", invoice.Number),
            invoice.TaxPointDate.HasValue ? Element("P_6", Date(invoice.TaxPointDate.Value)) : null,
            invoice.Period == null ? null : new XElement(Ns + "OkresFa", Element("P_6_Od", Date(invoice.Period.Start!.Value)), Element("P_6_Do", Date(invoice.Period.End!.Value))));
        foreach (Fa3TaxSummary tax in amounts.Taxes.OrderBy(item => Array.IndexOf(TaxSuffixes, item.FieldSuffix))) {
            fa.Add(Element("P_13_" + tax.FieldSuffix, Number(tax.TaxableAmount)));
            if (tax.TaxAmount.HasValue) fa.Add(Element("P_14_" + tax.FieldSuffix, Number(tax.TaxAmount.Value)));
            if (tax.TaxAmountInPln.HasValue) fa.Add(Element("P_14_" + tax.FieldSuffix + "W", Number(tax.TaxAmountInPln.Value)));
        }
        fa.Add(Element("P_15", Number(amounts.Total)), WriteAnnotations(options.Annotations), Element("RodzajFaktury", KindCode(options.Kind)));
        if (IsCorrection(options.Kind)) {
            fa.Add(Element("PrzyczynaKorekty", options.CorrectionReason));
            if (options.CorrectionTimingCode.HasValue) fa.Add(Element("TypKorekty", options.CorrectionTimingCode.Value.ToString(CultureInfo.InvariantCulture)));
            foreach (InvoiceReference reference in invoice.PrecedingInvoices) {
                options.CorrectionKsefNumbers.TryGetValue(reference.Number, out string? ksef);
                fa.Add(new XElement(Ns + "DaneFaKorygowanej", Element("DataWystFaKorygowanej", Date(reference.IssueDate!.Value)),
                    Element("NrFaKorygowanej", reference.Number), Element(ksef == null ? "NrKSeFN" : "NrKSeF", "1"), Element("NrKSeFFaKorygowanej", ksef)));
            }
            if (options.PreviousAdvanceOrSettlementTotal.HasValue) fa.Add(Element("P_15ZK", Number(options.PreviousAdvanceOrSettlementTotal.Value)));
        }
        foreach (Fa3AdvanceInvoiceReference reference in options.AdvanceInvoiceReferences)
            fa.Add(new XElement(Ns + "FakturaZaliczkowa", reference.KsefNumber == null ? Element("NrKSeFZN", "1") : null,
                Element("NrFaZaliczkowej", reference.Number), Element("NrKSeFFaZaliczkowej", reference.KsefNumber)));
        for (int index = 0; index < invoice.Lines.Count; index++) fa.Add(WriteLine(invoice.Lines[index], prepared.Calculation!.Lines[index].NetAmount, options, false));
        XElement? payment = WritePayment(invoice);
        if (payment != null) fa.Add(payment);
        if (options.Order != null) {
            var order = new XElement(Ns + "Zamowienie", Element("WartoscZamowienia", Number(options.Order.Total)));
            for (int index = 0; index < options.Order.Lines.Count; index++) {
                XElement row = WriteLine(options.Order.Lines[index], prepared.OrderCalculation!.Lines[index].NetAmount, options, true);
                decimal? vat = options.Order.DifferenceTaxAmounts?[index];
                if (vat.HasValue) row.Elements(Ns + "P_11NettoZ").Single().AddAfterSelf(Element("P_11VatZ", Number(vat.Value)));
                order.Add(row);
            }
            fa.Add(order);
        }
        var document = new XDocument(new XElement(Ns + "Faktura",
            new XElement(Ns + "Naglowek", new XElement(Ns + "KodFormularza", new XAttribute("kodSystemowy", "FA (3)"), new XAttribute("wersjaSchemy", "1-0E"), "FA"),
                Element("WariantFormularza", "3"), Element("DataWytworzeniaFa", options.CreatedAt.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture)), Element("SystemInfo", options.SystemInfo)),
            WriteParty(invoice.Seller, true, options.Kind), WriteParty(invoice.Buyer, false, options.Kind), fa));
        using var output = new InvoiceXmlOutputStream();
        using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
            Encoding = new UTF8Encoding(false), Indent = true, IndentChars = "  ", NewLineChars = "\n",
            NewLineHandling = NewLineHandling.Entitize, CloseOutput = false
        })) document.Save(writer);
        return output.ToArray();
    }

    private static string KindCode(Fa3InvoiceKind kind) => kind switch {
        Fa3InvoiceKind.TaxInvoice => "VAT", Fa3InvoiceKind.Correction => "KOR", Fa3InvoiceKind.Advance => "ZAL",
        Fa3InvoiceKind.Settlement => "ROZ", Fa3InvoiceKind.Simplified => "UPR", Fa3InvoiceKind.AdvanceCorrection => "KOR_ZAL",
        Fa3InvoiceKind.SettlementCorrection => "KOR_ROZ", _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };
    private static bool IsCorrection(Fa3InvoiceKind kind) => kind is Fa3InvoiceKind.Correction or Fa3InvoiceKind.AdvanceCorrection or Fa3InvoiceKind.SettlementCorrection;
    private static string Number(decimal value) => value.ToString("0.############################", CultureInfo.InvariantCulture);
    private static string Date(DateTime value) => value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    private static XElement? Element(string name, string? value) => value == null ? null : new XElement(Ns + name, value);
    private sealed class Prepared {
        internal IReadOnlyList<InvoiceDiagnostic> Diagnostics { get; set; } = Array.Empty<InvoiceDiagnostic>();
        internal InvoiceCalculation? Calculation { get; set; }
        internal InvoiceCalculation? OrderCalculation { get; set; }
        internal Fa3FiscalAmounts? Amounts { get; set; }
    }
}
