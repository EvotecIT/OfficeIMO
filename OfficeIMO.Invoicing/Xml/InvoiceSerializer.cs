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
    public static byte[] Write(Invoice invoice, InvoiceXmlOptions options) {
        if (invoice == null) throw new ArgumentNullException(nameof(invoice));
        if (options == null) throw new ArgumentNullException(nameof(options));
        InvoiceModelValidationResult validation = InvoiceModelValidator.ValidateForTarget(invoice, options);
        validation.ThrowIfInvalid();
        List<InvoiceDiagnostic> mapping = GetWriteDiagnostics(invoice, options);
        if (mapping.Any(item => item.Severity == InvoiceDiagnosticSeverity.Error))
            throw new InvalidDataException(string.Join(Environment.NewLine, mapping.Where(item => item.Severity == InvoiceDiagnosticSeverity.Error)
                .Select(item => item.Location + ": " + item.Message)));
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

    /// <summary>Validates the model and reports unsupported target mappings without writing or silently discarding information.</summary>
    public static IReadOnlyList<InvoiceDiagnostic> InspectTarget(Invoice invoice, InvoiceXmlOptions options) {
        if (invoice == null) throw new ArgumentNullException(nameof(invoice));
        if (options == null) throw new ArgumentNullException(nameof(options));
        InvoiceModelValidationResult validation = InvoiceModelValidator.ValidateForTarget(invoice, options);
        if (!validation.IsValid) return validation.Diagnostics;
        return GetWriteDiagnostics(invoice, options).AsReadOnly();
    }

    private static List<InvoiceDiagnostic> GetWriteDiagnostics(Invoice invoice, InvoiceXmlOptions options) {
        var diagnostics = new InvoiceDiagnosticBuffer();
        void Unsupported(string path, string text) => diagnostics.Add("INV-TARGET-UNSUPPORTED", text, path);
        void Projection(string path, string text) => diagnostics.Add("INV-TARGET-PROJECTION", text, path,
            options.ProjectionPolicy == InvoiceProjectionPolicy.AllowProfileDefinedDataLoss ? InvoiceDiagnosticSeverity.Warning : InvoiceDiagnosticSeverity.Error);
        if (invoice.Payments.Count != 0) {
            CheckPaymentProfile(invoice, options, Unsupported);
            CheckSingletonPaymentField(invoice, payment => payment.MeansCode, "MeansCode",
                "EN 16931 requires every payment-means occurrence to use the same payment means code", Unsupported);
            for (int index = 0; index < invoice.Payments.Count; index++) {
                InvoicePayment payment = invoice.Payments[index];
                string path = "Payments[" + index + "]";
                if (options.Syntax == InvoiceSyntax.Ubl && payment.CardNumber != null && string.IsNullOrWhiteSpace(payment.CardNetworkId))
                    Unsupported(path + ".CardNetworkId", "UBL requires a card network identifier when card account data is present.");
                if (options.Syntax == InvoiceSyntax.Cii && payment.DebitedAccount != null && !InvoiceBankAccountIdentity.IsValidIban(payment.DebitedAccount))
                    Unsupported(path + ".DebitedAccount", "The CII debtor-account target requires a valid IBAN; identifier '" + payment.DebitedAccount + "' cannot be relabeled as an IBAN.");
                InvoiceBankAccount? account = payment.Account;
                if (account == null) continue;
                bool validIban = InvoiceBankAccountIdentity.IsValidIban(account.Identifier);
                if (account.IsIban && !validIban)
                    Unsupported(path + ".Account", "An account marked as an IBAN must have a registered country format and valid checksum.");
                else if (options.Syntax == InvoiceSyntax.Ubl && !account.IsIban && validIban)
                    Unsupported(path + ".Account", "UBL cannot preserve an explicit proprietary-account classification for an identifier that is a valid IBAN.");
            }
            if (options.Syntax == InvoiceSyntax.Cii) {
                CheckSingletonPaymentField(invoice, payment => payment.Reference, "Reference", "CII has one invoice-level payment reference", Unsupported);
                CheckSingletonPaymentField(invoice, payment => payment.CreditorIdentifier, "CreditorIdentifier", "CII has one invoice-level creditor identifier", Unsupported);
                CheckSingletonPaymentField(invoice, payment => payment.MandateReference, "MandateReference", "CII has one invoice-level direct-debit mandate reference", Unsupported);
            } else {
                CheckSingletonPaymentField(invoice, payment => payment.CreditorIdentifier, "CreditorIdentifier", "UBL carries one seller-level SEPA creditor identifier", Unsupported);
            }
        }
        if (options.Syntax == InvoiceSyntax.Ubl && invoice.TypeCode != "380" && invoice.TypeCode != "381" && invoice.TypeCode != "384" && invoice.TypeCode != "389")
            Unsupported("TypeCode", "UBL authoring supports invoice codes 380, 384, 389 and credit note code 381.");
        if (options.Syntax == InvoiceSyntax.Ubl) {
            if (invoice.Buyer.Identifiers.Count > 1)
                Unsupported("Buyer.Identifiers", "The EN 16931 UBL mapping permits at most one buyer identifier; CII can preserve multiple buyer identifiers.");
            foreach (InvoiceNote note in invoice.Notes)
                if (note.SubjectCode == null && InvoiceNote.HasEncodedSubject(note.Text))
                    Unsupported("Notes", "An unclassified note begins with a reserved UBL subject prefix and cannot be represented without changing its meaning.");
            if (invoice.TypeCode == "381" && invoice.DueDate.HasValue) Unsupported("DueDate", "UBL credit note due-date mapping is not supported.");
            if (invoice.TypeCode == "381" && invoice.ProjectReference != null) Unsupported("ProjectReference", "UBL credit note project reference mapping is not supported.");
            if (string.IsNullOrWhiteSpace(invoice.PurchaseOrderReference) && invoice.SalesOrderReference != null) Unsupported("SalesOrderReference", "UBL requires a purchase order reference alongside a sales order reference.");
            if (invoice.Seller.Identifiers.Any(id => id.SchemeId == "SEPA")) Unsupported("Seller.Identifiers", "Use Payment.CreditorIdentifier for the reserved SEPA creditor identifier.");
            if (invoice.Buyer.Identifiers.Any(id => id.SchemeId == "SEPA")) Unsupported("Buyer.Identifiers", "The reserved SEPA creditor identifier belongs to the seller payment details, not the buyer.");
        }
        if (options.Profile is (InvoiceProfile.XRechnung or InvoiceProfile.PeppolBis) && string.IsNullOrWhiteSpace(invoice.BusinessProcessId))
            Unsupported("BusinessProcessId", "XRechnung/Peppol output requires an explicit business process identifier.");
        if (options.Profile == InvoiceProfile.XRechnung && string.IsNullOrWhiteSpace(invoice.BuyerReference))
            Unsupported("BuyerReference", "XRechnung output requires a buyer routing reference.");
        CheckExemptionConflicts(invoice, Unsupported);
        CheckFacturXProjection(invoice, options, Projection);
        CheckTaxRegistrations(invoice.Seller, "Seller", options, Unsupported);
        CheckTaxRegistrations(invoice.Buyer, "Buyer", options, Unsupported);
        if (invoice.Payee != null) CheckTaxRegistrations(invoice.Payee, "Payee", options, Unsupported);
        if (invoice.TaxRepresentative != null) CheckTaxRegistrations(invoice.TaxRepresentative, "TaxRepresentative", options, Unsupported);
        // The EN semantic payee and tax-representative groups are intentionally narrower than seller/buyer.
        InvoicePartyMapping.Check(invoice.Payee, "Payee", Unsupported);
        InvoicePartyMapping.Check(invoice.TaxRepresentative, "TaxRepresentative", Unsupported);
        InvoicePartyMapping.Check(invoice.Buyer, "Buyer", Unsupported);
        return diagnostics.ToList();
    }
    private static void CheckTaxRegistrations(InvoiceParty party, string role, InvoiceXmlOptions options, Action<string, string> unsupported) {
        int vatCount = 0, otherCount = 0;
        for (int index = 0; index < party.TaxRegistrations.Count; index++) {
            InvoiceTaxRegistration registration = party.TaxRegistrations[index];
            string path = role + ".TaxRegistrations[" + index + "]";
            if (registration.SchemeId == InvoiceTaxRegistration.VatScheme) {
                if (++vatCount > 1) unsupported(path, "The target profile permits at most one VAT registration for this party; the additional occurrence remains available in the source model.");
                continue;
            }
            otherCount++;
            if (role != "Seller") {
                unsupported(path, "The target profile has no semantic field for a non-VAT " + role + " tax registration with scheme '" + registration.SchemeId + "'.");
            } else if (otherCount > 1) {
                unsupported(path, "The target profile permits at most one non-VAT seller tax registration; this occurrence uses scheme '" + registration.SchemeId + "'.");
            } else if (options.Syntax == InvoiceSyntax.Cii && registration.SchemeId != InvoiceTaxRegistration.TaxScheme) {
                unsupported(path, "CII EN 16931 maps the seller fiscal registration through scheme 'FC' and cannot preserve source scheme '" + registration.SchemeId + "' without relabeling it.");
            }
        }
    }
    private static void CheckSingletonPaymentField(Invoice invoice, Func<InvoicePayment, string?> selector, string field, string targetContract,
        Action<string, string> unsupported) {
        string? retained = null;
        for (int index = 0; index < invoice.Payments.Count; index++) {
            string? value = selector(invoice.Payments[index]);
            if (value == null) continue;
            if (retained == null) retained = value;
            else if (!string.Equals(retained, value, StringComparison.Ordinal))
                unsupported("Payments[" + index + "]." + field, targetContract + "; conflicting value '" + value + "' cannot be combined with '" + retained + "'.");
        }
    }
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
    private static string? FirstPaymentValue(Invoice invoice, Func<InvoicePayment, string?> selector) =>
        invoice.Payments.Select(selector).FirstOrDefault(value => value != null);
}
