using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class Fa3InvoiceWriter {
    private static XElement WriteAnnotations(Fa3InvoiceAnnotations value) {
        var exemption = new XElement(Ns + "Zwolnienie", value.ExemptionLegalBasis == null
            ? Element("P_19N", "1") : null);
        if (value.ExemptionLegalBasis != null) exemption.Add(Element("P_19", "1"),
            Element(value.ExemptionBasisKind == Fa3ExemptionBasisKind.NationalProvision ? "P_19A" : value.ExemptionBasisKind == Fa3ExemptionBasisKind.EuDirective ? "P_19B" : "P_19C", value.ExemptionLegalBasis));
        var margin = new XElement(Ns + "PMarzy", value.MarginProcedure == Fa3MarginProcedure.None
            ? Element("P_PMarzyN", "1") : Element("P_PMarzy", "1"));
        if (value.MarginProcedure != Fa3MarginProcedure.None) margin.Add(Element(value.MarginProcedure switch {
            Fa3MarginProcedure.Travel => "P_PMarzy_2", Fa3MarginProcedure.UsedGoods => "P_PMarzy_3_1",
            Fa3MarginProcedure.Art => "P_PMarzy_3_2", _ => "P_PMarzy_3_3"
        }, "1"));
        return new XElement(Ns + "Adnotacje", Element("P_16", value.CashAccounting ? "1" : "2"),
            Element("P_17", value.SelfBilling ? "1" : "2"), Element("P_18", value.ReverseCharge ? "1" : "2"), Element("P_18A", value.SplitPayment ? "1" : "2"),
            exemption, new XElement(Ns + "NoweSrodkiTransportu", Element("P_22N", "1")),
            Element("P_23", value.TriangularTransaction ? "1" : "2"), margin);
    }

    private static XElement WriteParty(InvoiceParty party, bool seller, Fa3InvoiceKind kind) {
        string? nip = Nip(party);
        var identity = new XElement(Ns + "DaneIdentyfikacyjne", nip == null ? Element("BrakID", "1") : Element("NIP", nip),
            string.IsNullOrEmpty(party.Name) && kind == Fa3InvoiceKind.Simplified && !seller ? null : Element("Nazwa", party.Name));
        var result = new XElement(Ns + (seller ? "Podmiot1" : "Podmiot2"), identity);
        if (!string.IsNullOrEmpty(party.Address.Line1)) result.Add(new XElement(Ns + "Adres",
            Element("KodKraju", party.Address.CountryCode), Element("AdresL1", party.Address.Line1), Element("AdresL2", party.Address.Line2)));
        if (party.Contact?.Email != null || party.Contact?.Telephone != null)
            result.Add(new XElement(Ns + "DaneKontaktowe", Element("Email", party.Contact?.Email), Element("Telefon", party.Contact?.Telephone)));
        if (!seller) result.Add(Element("JST", "2"), Element("GV", "2"));
        return result;
    }

    private static XElement WriteLine(InvoiceLine line, decimal net, Fa3InvoiceWriteOptions options, bool order) {
        var row = new XElement(Ns + (order ? "ZamowienieWiersz" : "FaWiersz"), Element(order ? "NrWierszaZam" : "NrWierszaFa", line.Id));
        row.Add(Element(order ? "P_7Z" : "P_7", line.Name), Element(order ? "IndeksZ" : "Indeks", line.SellerItemIdentifier),
            Element(order ? "P_8AZ" : "P_8A", UnitLabel(line.UnitCode, options)), Element(order ? "P_8BZ" : "P_8B", Number(line.Quantity)),
            Element(order ? "P_9AZ" : "P_9A", Number(line.UnitPrice)), Element(order ? "P_11NettoZ" : "P_11", Number(net)));
        if (options.Annotations.MarginProcedure == Fa3MarginProcedure.None) row.Add(Element(order ? "P_12Z" : "P_12", TaxLabel(line, options)));
        return row;
    }

    private static XElement? WritePayment(Invoice invoice) {
        if (!invoice.DueDate.HasValue && invoice.Payments.Count == 0) return null;
        var result = new XElement(Ns + "Platnosc");
        if (invoice.DueDate.HasValue) result.Add(new XElement(Ns + "TerminPlatnosci", Element("Termin", Date(invoice.DueDate.Value))));
        if (invoice.Payments.Count != 0) result.Add(Element("FormaPlatnosci", PaymentForm(invoice.Payments[0].MeansCode)));
        foreach (InvoicePayment payment in invoice.Payments) if (payment.Account != null)
            result.Add(new XElement(Ns + "RachunekBankowy", Element("NrRB", payment.Account.Identifier), Element("SWIFT", payment.Account.ProviderIdentifier)));
        return result;
    }
    private static string? PaymentForm(string means) => means switch { "10" => "1", "48" => "2", "20" => "4", "30" or "58" => "6", _ => null };
    private static string? Nip(InvoiceParty party) {
        InvoiceTaxRegistration? tax = party.TaxRegistrations.FirstOrDefault();
        if (tax == null) return null;
        return tax.Kind == InvoiceTaxRegistrationKind.Vat && tax.Identifier.StartsWith("PL", StringComparison.Ordinal)
            ? tax.Identifier.Substring(2) : tax.Identifier;
    }
    private static string? UnitLabel(string code, Fa3InvoiceWriteOptions options) {
        if (options.UnitLabels.TryGetValue(code, out string? label)) return label;
        return code switch { "C62" => "szt.", "KGM" => "kg", "GRM" => "g", "LTR" => "l", "MTR" => "m", "MTK" => "m2", "MTQ" => "m3", "HUR" => "godz.", _ => null };
    }
    private static string? TaxLabel(InvoiceLine line, Fa3InvoiceWriteOptions options) {
        if (options.LineTaxLabels.TryGetValue(line.Id, out string? label)) return label;
        return line.Tax.Code switch {
            "S" => line.Tax.Rate.HasValue ? Number(line.Tax.Rate.Value) : null,
            "Z" => "0 KR", "K" => "0 WDT", "G" => "0 EX", "E" => "zw", "AE" => "oo", _ => null
        };
    }
    private static string? TaxSuffix(string? label) => label switch {
        "23" or "22" => "1", "8" or "7" => "2", "5" => "3", "4" or "3" => "4", "0 KR" => "6_1",
        "0 WDT" => "6_2", "0 EX" => "6_3", "zw" => "7", "np I" => "8", "np II" => "9", "oo" => "10", _ => null
    };
}
