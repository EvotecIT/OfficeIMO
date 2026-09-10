using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

/// <summary>Native deterministic CII/UBL serialization over the semantic invoice model.</summary>
public static partial class InvoiceSerializer {
    private static readonly XNamespace Ram = InvoiceXml.Ram;
    private static readonly XNamespace Rsm = InvoiceXml.Rsm;
    private static readonly XNamespace Udt = "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100";
    private static readonly XNamespace Qdt = "urn:un:unece:uncefact:data:standard:QualifiedDataType:100";
    private static readonly XNamespace Cbc = InvoiceXml.Cbc;
    private static readonly XNamespace Cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";

    /// <summary>Validates the model and writes UTF-8 XML without a BOM. Identical models and options produce identical bytes.</summary>
    public static byte[] Write(Invoice invoice, InvoiceXmlOptions? options = null) {
        if (invoice == null) throw new ArgumentNullException(nameof(invoice));
        options = options ?? new InvoiceXmlOptions();
        InvoiceModelValidationResult validation = InvoiceModelValidator.Validate(invoice);
        validation.ThrowIfInvalid();
        List<InvoiceDiagnostic> mapping = GetWriteDiagnostics(invoice, options);
        if (mapping.Count != 0) throw new InvalidDataException(string.Join(Environment.NewLine, mapping.Select(item => item.Location + ": " + item.Message)));
        XDocument document = options.Syntax == InvoiceSyntax.Cii
            ? WriteCii(invoice, validation.Calculation!, options)
            : WriteUbl(invoice, validation.Calculation!, options);
        using var output = new InvoiceXmlOutputStream();
        using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
            Encoding = new UTF8Encoding(false), Indent = true, IndentChars = "  ", NewLineChars = "\n",
            NewLineHandling = NewLineHandling.Entitize, CloseOutput = false
        })) document.Save(writer);
        return output.ToArray();
    }

    /// <summary>Returns unsupported target mappings without writing or silently discarding information.</summary>
    public static IReadOnlyList<InvoiceDiagnostic> InspectTarget(Invoice invoice, InvoiceXmlOptions options) {
        if (invoice == null) throw new ArgumentNullException(nameof(invoice));
        if (options == null) throw new ArgumentNullException(nameof(options));
        return GetWriteDiagnostics(invoice, options).AsReadOnly();
    }

    private static List<InvoiceDiagnostic> GetWriteDiagnostics(Invoice invoice, InvoiceXmlOptions options) {
        var diagnostics = new List<InvoiceDiagnostic>();
        void Unsupported(string path, string text) => diagnostics.Add(new InvoiceDiagnostic("INV-TARGET-UNSUPPORTED", text, path));
        if (options.Syntax == InvoiceSyntax.Ubl && invoice.TypeCode != "380" && invoice.TypeCode != "381" && invoice.TypeCode != "384" && invoice.TypeCode != "389")
            Unsupported("TypeCode", "UBL authoring supports invoice codes 380, 384, 389 and credit note code 381.");
        if (options.Syntax == InvoiceSyntax.Ubl) {
            foreach (InvoiceNote note in invoice.Notes)
                if (note.SubjectCode == null && InvoiceNote.HasEncodedSubject(note.Text))
                    Unsupported("Notes", "An unclassified note begins with a reserved UBL subject prefix and cannot be represented without changing its meaning.");
            if (invoice.TypeCode == "381" && invoice.DueDate.HasValue) Unsupported("DueDate", "UBL credit note due-date mapping is not supported.");
            if (invoice.TypeCode == "381" && invoice.ProjectReference != null) Unsupported("ProjectReference", "UBL credit note project reference mapping is not supported.");
            if (string.IsNullOrWhiteSpace(invoice.PurchaseOrderReference) && invoice.SalesOrderReference != null) Unsupported("SalesOrderReference", "UBL requires a purchase order reference alongside a sales order reference.");
            if (invoice.Seller.Identifiers.Any(id => id.SchemeId == "SEPA")) Unsupported("Seller.Identifiers", "Use Payment.CreditorIdentifier for the reserved SEPA creditor identifier.");
            if (invoice.Buyer.Identifiers.Any(id => id.SchemeId == "SEPA")) Unsupported("Buyer.Identifiers", "The reserved SEPA creditor identifier belongs to the seller payment details, not the buyer.");
        }
        if (options.Profile != InvoiceProfile.En16931 && string.IsNullOrWhiteSpace(invoice.BusinessProcessId))
            Unsupported("BusinessProcessId", "XRechnung/Peppol output requires an explicit business process identifier.");
        if (options.Profile == InvoiceProfile.XRechnung && string.IsNullOrWhiteSpace(invoice.BuyerReference))
            Unsupported("BuyerReference", "XRechnung output requires a buyer routing reference.");
        // The EN semantic payee and tax-representative groups are intentionally narrower than seller/buyer.
        CheckParty(invoice.Payee, "Payee", true);
        CheckParty(invoice.TaxRepresentative, "TaxRepresentative", false);
        if (invoice.Buyer?.TaxRegistration != null) Unsupported("Buyer.TaxRegistration", "EN 16931 has no buyer non-VAT tax registration term.");
        if (invoice.Buyer?.LegalInformation != null) Unsupported("Buyer.LegalInformation", "EN 16931 has no buyer additional legal information term.");
        return diagnostics;

        void CheckParty(InvoiceParty? party, string path, bool payee) {
            if (party == null) return;
            if (party.TradingName != null || party.LegalInformation != null || party.Contact != null || party.ElectronicAddress != null || party.TaxRegistration != null)
                Unsupported(path, "This party role cannot carry trading name, legal information, contact, electronic address or non-VAT tax registration in the supported mapping.");
            if (payee && (party.VatIdentifier != null || HasAddress(party.Address))) Unsupported(path, "The payee group does not carry VAT or postal-address information.");
            if (!payee && (party.Identifiers.Count != 0 || party.LegalRegistration != null)) Unsupported(path, "The tax representative group does not carry business or legal registration identifiers.");
            if (payee && party.Identifiers.Count > 1) Unsupported(path + ".Identifiers", "The payee group permits one identifier.");
        }
    }

    private static bool HasAddress(InvoiceAddress? address) => address != null &&
        (!string.IsNullOrEmpty(address.CountryCode) || address.Line1 != null || address.Line2 != null || address.Line3 != null || address.City != null || address.PostCode != null || address.Subdivision != null);
    private static string Number(decimal value) => value.ToString("0.############################", CultureInfo.InvariantCulture);
    private static string Amount(decimal value) => value.ToString("0.00", CultureInfo.InvariantCulture);
    private static XElement? Text(XName name, string? value) => value == null ? null : new XElement(name, value);
    private static XElement? CiiDate(string name, DateTime? value, bool qualified = false) => value.HasValue
        ? new XElement(Ram + name, new XElement((qualified ? Qdt : Udt) + "DateTimeString", new XAttribute("format", "102"), value.Value.ToString("yyyyMMdd", CultureInfo.InvariantCulture))) : null;
    private static XElement? UblDate(string name, DateTime? value) => value.HasValue ? new XElement(Cbc + name, value.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)) : null;
    private static XElement UblAmount(string name, decimal value, string currency) => new XElement(Cbc + name, new XAttribute("currencyID", currency), Amount(value));
    private static XElement CiiAmount(string name, decimal value) => new XElement(Ram + name, Amount(value));
    private static XElement? Identifier(XName name, InvoiceIdentifier? value, string schemeAttribute = "schemeID") => value == null ? null :
        new XElement(name, value.SchemeId == null ? null : new XAttribute(schemeAttribute, value.SchemeId), value.Value);
}
