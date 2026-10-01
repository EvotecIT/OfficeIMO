using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

/// <summary>Bounded FA(3) reading with exact national amounts and explicit common-model projection limits.</summary>
public static partial class Fa3InvoiceReader {
    /// <summary>Namespace of the FA(3) schema published on 25 June 2025.</summary>
    public const string NamespaceUri = "http://crd.gov.pl/wzor/2025/06/25/13775/";
    private static readonly XNamespace Ns = NamespaceUri;
    private static readonly string[] TaxSuffixes = { "1", "2", "3", "4", "5", "6_1", "6_2", "6_3", "7", "8", "9", "10", "11" };

    /// <summary>Reads a defensive snapshot, rejecting DTDs, ambiguous scalar fields and inputs over 16 MiB or depth 128.</summary>
    /// <remarks>National VAT buckets and P_15 are retained verbatim. No rounding, exchange rate or advance settlement is invented.</remarks>
    public static Fa3InvoiceReadResult Read(byte[] xml) {
        if (xml == null) throw new ArgumentNullException(nameof(xml));
        if (xml.Length == 0 || xml.Length > InvoiceProfileDeclaration.MaximumXmlBytes)
            throw new InvalidDataException("Invoice XML must contain between 1 byte and 16 MiB.");
        byte[] snapshot = (byte[])xml.Clone();
        XElement root = InvoiceXml.Parse(snapshot).Root ?? throw new InvalidDataException("Missing FA(3) root.");
        if (root.Name != Ns + "Faktura") throw new InvalidDataException("Expected the pinned FA(3) Faktura namespace.");
        var c = new InvoiceXmlReadContext(); c.Consume(root);
        var projection = new InvoiceDiagnosticBuffer();
        XElement header = c.Require(root, Ns + "Naglowek");
        XElement form = c.Require(header, Ns + "KodFormularza");
        if (c.Value(form) != "FA" || c.Attribute(form, "kodSystemowy") != "FA (3)" ||
            c.Attribute(form, "wersjaSchemy") != "1-0E" || c.Required(header, Ns + "WariantFormularza") != "3")
            throw new InvalidDataException("FA(3) form declaration is inconsistent with FA (3), schema 1-0E, variant 3.");
        string timestamp = c.Required(header, Ns + "DataWytworzeniaFa");
        if (!(timestamp.EndsWith("Z", StringComparison.Ordinal) ||
            timestamp.Length >= 6 && (timestamp[timestamp.Length - 6] == '+' || timestamp[timestamp.Length - 6] == '-') && timestamp[timestamp.Length - 3] == ':') ||
            !DateTimeOffset.TryParseExact(timestamp, new[] { "yyyy-MM-dd'T'HH:mm:ssK", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK" },
                CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTimeOffset created))
            throw new InvalidDataException("FA(3) creation timestamp requires an explicit timezone.");
        string? system = c.Text(header, Ns + "SystemInfo");
        XElement fa = c.Require(root, Ns + "Fa");
        Fa3InvoiceKind kind = ReadKind(c.Required(fa, Ns + "RodzajFaktury"));
        var invoice = new Invoice {
            Number = c.Required(fa, Ns + "P_2"), IssueDate = c.Date(c.Require(fa, Ns + "P_1"))!.Value,
            Currency = c.Required(fa, Ns + "KodWaluty"),
            TypeCode = kind == Fa3InvoiceKind.Correction || kind == Fa3InvoiceKind.AdvanceCorrection || kind == Fa3InvoiceKind.SettlementCorrection ? "384" :
                kind == Fa3InvoiceKind.Advance ? "386" : "380",
            Seller = ReadParty(c, c.Require(root, Ns + "Podmiot1"), true),
            Buyer = ReadParty(c, c.Require(root, Ns + "Podmiot2"), false),
            TaxPointDate = c.Date(c.Child(fa, Ns + "P_6"))
        };
        string? place = c.Text(fa, Ns + "P_1M");
        XElement? period = c.Child(fa, Ns + "OkresFa");
        if (period != null) invoice.Period = new InvoicePeriod {
            Start = c.Date(c.Child(period, Ns + "P_6_Od")), End = c.Date(c.Child(period, Ns + "P_6_Do"))
        };
        decimal total = c.RequiredDecimal(c.Require(fa, Ns + "P_15"));
        var taxes = new List<Fa3TaxSummary>();
        foreach (string suffix in TaxSuffixes) {
            decimal? basis = c.Decimal(fa, Ns + ("P_13_" + suffix));
            if (basis.HasValue) c.AddTo(taxes, new Fa3TaxSummary(suffix, basis.Value,
                c.Decimal(fa, Ns + ("P_14_" + suffix)), c.Decimal(fa, Ns + ("P_14_" + suffix + "W"))));
        }
        // A national bucket can combine historical rates and has different semantics for advances/margin procedures.
        // Keep exact declarations in TaxSummaries; do not guess EN breakdowns or replace P_15 by a computed sum.
        projection.Add("FA3-NATIONAL-TOTALS", "National VAT buckets and P_15 require an explicit fiscal mapping before EN conversion.", "Fa.Totals");
        if (kind != Fa3InvoiceKind.TaxInvoice && kind != Fa3InvoiceKind.Simplified)
            projection.Add("FA3-DOCUMENT-KIND", "Signed corrections, advances and settlements retain their national semantics; the common type code is a presentation hint.", "Fa.RodzajFaktury");
        var lines = new List<Fa3InvoiceLineData>();
        foreach (XElement line in c.Children(fa, Ns + "FaWiersz")) {
            if (lines.Count >= 10_000) throw new InvalidDataException("FA(3) exceeds 10,000 invoice rows.");
            ReadLine(c, line, invoice, lines, projection, kind == Fa3InvoiceKind.TaxInvoice || kind == Fa3InvoiceKind.Simplified);
        }
        ReadPayment(c, fa, invoice, projection);
        new InvoiceModelLimits().Check(invoice);
        return new Fa3InvoiceReadResult(snapshot, invoice, kind, created, system, place, total,
            taxes.AsReadOnly(), lines.AsReadOnly(), c.Finish(root), projection.ToList().AsReadOnly());
    }

    private static Fa3InvoiceKind ReadKind(string code) => code switch {
        "VAT" => Fa3InvoiceKind.TaxInvoice, "KOR" => Fa3InvoiceKind.Correction, "ZAL" => Fa3InvoiceKind.Advance,
        "ROZ" => Fa3InvoiceKind.Settlement, "UPR" => Fa3InvoiceKind.Simplified,
        "KOR_ZAL" => Fa3InvoiceKind.AdvanceCorrection, "KOR_ROZ" => Fa3InvoiceKind.SettlementCorrection,
        _ => throw new InvalidDataException("Unknown FA(3) document kind.")
    };
}
