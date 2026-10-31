using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using Shared = OfficeIMO.Internal.Invoicing;

namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    // These profiles intentionally omit EN 16931 line, payment and party fields.
    // Select from the XML, then independently check that XMP agrees with it.
    private static bool TryGetReducedInvoiceProfile(PdfCreateOptions options, out Shared.InvoiceProfile profile) {
        profile = default;
        PdfEmbeddedFile? file = options.EmbeddedFiles.FirstOrDefault(item => item.FileName == "factur-x.xml");
        return file != null && TryReadCiiProfileContext(file, out string? context, out _) &&
            Shared.InvoiceProfiles.TryFromGuidelineId(context, out profile) &&
            profile is Shared.InvoiceProfile.Minimum or Shared.InvoiceProfile.BasicWithoutLines or Shared.InvoiceProfile.Basic;
    }

    private static void AddReducedInvoiceRequirements(List<PdfComplianceRequirement> requirements, PdfCreateOptions options, Shared.InvoiceProfile profile) {
        requirements.Add(BuildElectronicInvoiceXmlAttachmentRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlProfileContextRequirement(options));
        requirements.Add(BuildElectronicInvoiceProfileConsistencyRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlDocumentHeaderRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlDocumentTypeCodeRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlDateFormatRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTradeTransactionRequirement(options));
        requirements.Add(BuildReducedInvoiceFieldsRequirement(options, profile));
        if (profile != Shared.InvoiceProfile.Minimum) {
            requirements.Add(BuildElectronicInvoiceXmlPartyIdentificationRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlCountryCodeRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlTaxBreakdownRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlTaxCategoryCodeRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlTaxCategoryRateRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlTaxCategoryAmountRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlTaxExemptionReasonRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlTaxPartyIdentifierRequirement(options));
        }
        if (profile == Shared.InvoiceProfile.Basic) {
            requirements.Add(BuildElectronicInvoiceXmlLineItemRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlUnitCodeRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlLinePricingRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlLineAmountConsistencyRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlLineTaxRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlTaxCategoryConsistencyRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlTaxTotalConsistencyRequirement(options));
            requirements.Add(BuildElectronicInvoiceXmlAmountConsistencyRequirement(options));
        }
        requirements.Add(BuildElectronicInvoiceXmlAttachmentParamsRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmpRequirement(options));
        requirements.Add(new PdfComplianceRequirement("mustang-validation", "Mustang validation evidence",
            PdfComplianceRequirementStatus.Unsupported,
            "Run Mustang on the exact XML and saved PDF to validate the declared reduced Factur-X profile."));
    }

    // Lightweight structural and arithmetic readiness, not a replacement for the
    // profile-specific XSD/Schematron authority gate in Invoicing.Validation.
    private static PdfComplianceRequirement BuildReducedInvoiceFieldsRequirement(PdfCreateOptions options, Shared.InvoiceProfile profile) {
        var diagnostics = new List<string>();
        try {
            PdfEmbeddedFile file = options.EmbeddedFiles.First(item => item.FileName == "factur-x.xml");
            XDocument document = Shared.InvoiceXml.Parse(file.DataSnapshot,
                Shared.InvoiceXml.MaximumNodes, Shared.InvoiceXml.MaximumAttributes);
            XNamespace ram = Shared.InvoiceXml.Ram;
            XElement root = document.Root!;
            XElement? Field(XElement? parent, XName name) => parent == null ? null : Shared.InvoiceXml.Unique(parent, name);
            string? Text(XElement? parent, XName name) {
                XElement? element = Field(parent, name);
                return element == null ? null : Shared.InvoiceXml.Scalar(element);
            }
            XElement? trade = Field(root, Shared.InvoiceXml.Rsm + "SupplyChainTradeTransaction");
            XElement? agreement = Field(trade, ram + "ApplicableHeaderTradeAgreement");
            foreach (string role in new[] { "SellerTradeParty", "BuyerTradeParty" }) {
                XElement? party = Field(agreement, ram + role);
                if (string.IsNullOrWhiteSpace(Text(party, ram + "Name"))) diagnostics.Add(role + " Name is required.");
                string? country = Text(Field(party, ram + "PostalTradeAddress"), ram + "CountryID");
                if (country != null && !ElectronicInvoiceCountryCodes.Contains(country, StringComparer.Ordinal))
                    diagnostics.Add(role + " CountryID must be an ISO 3166-1 alpha-2 code.");
            }
            XElement? settlement = Field(trade, ram + "ApplicableHeaderTradeSettlement");
            string? currency = Text(settlement, ram + "InvoiceCurrencyCode");
            if (currency == null || !IsKnownElectronicInvoiceCurrencyCode(currency)) diagnostics.Add("InvoiceCurrencyCode is required and must be an ISO 4217 code.");
            XElement? totals = Field(settlement, ram + "SpecifiedTradeSettlementHeaderMonetarySummation");
            string? accountingCurrency = Text(settlement, ram + "TaxCurrencyCode");
            if (accountingCurrency != null && (!IsKnownElectronicInvoiceCurrencyCode(accountingCurrency) || accountingCurrency == currency))
                diagnostics.Add("TaxCurrencyCode must be a distinct ISO 4217 accounting currency.");
            XElement[] taxTotals = totals?.Elements(ram + "TaxTotalAmount").ToArray() ?? Array.Empty<XElement>();
            foreach (var group in taxTotals.GroupBy(element => (string?)element.Attribute("currencyID"), StringComparer.Ordinal)) {
                if (group.Count() > 1) diagnostics.Add("Duplicate TaxTotalAmount for currency " + group.Key + ".");
                if (group.Key == null || group.Key != currency && group.Key != accountingCurrency)
                    diagnostics.Add("TaxTotalAmount currencyID must match InvoiceCurrencyCode or TaxCurrencyCode.");
            }
            XElement? taxTotal = taxTotals.FirstOrDefault(element => (string?)element.Attribute("currencyID") == currency);
            decimal? ParseAmount(XElement? element, string name, bool required) {
                string? value = element == null ? null : Shared.InvoiceXml.Scalar(element);
                if (value == null && !required) return 0m;
                if (value == null || !decimal.TryParse(value, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture, out decimal number) || decimal.Round(number, 2) != number) {
                    diagnostics.Add(name + " must be a decimal amount with at most two fractional digits.");
                    return null;
                }
                return number;
            }
            decimal? Amount(string name, bool required = false) =>
                ParseAmount(name == "TaxTotalAmount" ? taxTotal : Field(totals, ram + name), name, required);
            foreach (XElement otherTax in taxTotals.Where(element => !ReferenceEquals(element, taxTotal)))
                ParseAmount(otherTax, "TaxTotalAmount", true);
            decimal? basis = Amount("TaxBasisTotalAmount", true), tax = Amount("TaxTotalAmount"),
                grand = Amount("GrandTotalAmount", true), payable = Amount("DuePayableAmount", true), prepaid = Amount("TotalPrepaidAmount");
            if ((profile != Shared.InvoiceProfile.Minimum || taxTotal != null) &&
                basis.HasValue && tax.HasValue && grand.HasValue && basis.Value + tax.Value != grand.Value)
                diagnostics.Add("GrandTotalAmount must equal TaxBasisTotalAmount plus TaxTotalAmount.");
            if (profile != Shared.InvoiceProfile.Minimum && grand.HasValue && prepaid.HasValue && payable.HasValue && grand.Value - prepaid.Value != payable.Value)
                diagnostics.Add("DuePayableAmount must equal GrandTotalAmount minus TotalPrepaidAmount.");
            if (profile != Shared.InvoiceProfile.Minimum) {
                decimal? line = Amount("LineTotalAmount", true), charge = Amount("ChargeTotalAmount"), allowance = Amount("AllowanceTotalAmount");
                if (line.HasValue && charge.HasValue && allowance.HasValue && basis.HasValue && line.Value + charge.Value - allowance.Value != basis.Value)
                    diagnostics.Add("TaxBasisTotalAmount must equal LineTotalAmount plus charges minus allowances.");
            }
            if (profile != Shared.InvoiceProfile.Basic && trade?.Elements(ram + "IncludedSupplyChainTradeLineItem").Any() == true)
                diagnostics.Add("MINIMUM and BASIC WL cannot contain invoice lines.");
        } catch (Exception exception) when (exception is InvalidDataException || exception is XmlException || exception is OverflowException) {
            diagnostics.Add(exception.Message);
        }
        return new PdfComplianceRequirement("einvoice-xml-reduced-profile", "Reduced Factur-X profile essentials",
            diagnostics.Count == 0 ? PdfComplianceRequirementStatus.Satisfied : PdfComplianceRequirementStatus.Missing,
            diagnostics.Count == 0 ? "The reduced profile retains its required parties, currency and consistent aggregate amounts. Full profile validation requires Mustang."
                : string.Join(" ", diagnostics));
    }
}
