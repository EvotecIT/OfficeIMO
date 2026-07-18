namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static readonly string[] ElectronicInvoiceCurrencyCodes = {
        "AED", "AFN", "ALL", "AMD", "ANG", "AOA", "ARS", "AUD", "AWG", "AZN",
        "BAM", "BBD", "BDT", "BGN", "BHD", "BIF", "BMD", "BND", "BOB", "BOV",
        "BRL", "BSD", "BTN", "BWP", "BYN", "BZD", "CAD", "CDF", "CHE", "CHF",
        "CHW", "CLF", "CLP", "CNY", "COP", "COU", "CRC", "CUC", "CUP", "CVE",
        "CZK", "DJF", "DKK", "DOP", "DZD", "EGP", "ERN", "ETB", "EUR", "FJD",
        "FKP", "GBP", "GEL", "GHS", "GIP", "GMD", "GNF", "GTQ", "GYD", "HKD",
        "HNL", "HTG", "HUF", "IDR", "ILS", "INR", "IQD", "IRR", "ISK", "JMD",
        "JOD", "JPY", "KES", "KGS", "KHR", "KMF", "KPW", "KRW", "KWD", "KYD",
        "KZT", "LAK", "LBP", "LKR", "LRD", "LSL", "LYD", "MAD", "MDL", "MGA",
        "MKD", "MMK", "MNT", "MOP", "MRU", "MUR", "MVR", "MWK", "MXN", "MXV",
        "MYR", "MZN", "NAD", "NGN", "NIO", "NOK", "NPR", "NZD", "OMR", "PAB",
        "PEN", "PGK", "PHP", "PKR", "PLN", "PYG", "QAR", "RON", "RSD", "RUB",
        "RWF", "SAR", "SBD", "SCR", "SDG", "SEK", "SGD", "SHP", "SLE", "SOS",
        "SRD", "SSP", "STN", "SVC", "SYP", "SZL", "THB", "TJS", "TMT", "TND",
        "TOP", "TRY", "TTD", "TWD", "TZS", "UAH", "UGX", "USD", "USN", "UYI",
        "UYU", "UYW", "UZS", "VED", "VES", "VND", "VUV", "WST", "XAF", "XAG",
        "XAU", "XBA", "XBB", "XBC", "XBD", "XCD", "XDR", "XOF", "XPD", "XPF",
        "XPT", "XSU", "XTS", "XUA", "YER", "ZAR", "ZMW", "ZWG", "ZWL"
    };

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlCurrencyCodeRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadCurrencyConsistency(file, out PdfCiiCurrencyConsistencyEvidence? evidence, out string? currencyDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-currency-code",
                    "EN 16931 XML currency code",
                    PdfComplianceRequirementStatus.Missing,
                    currencyDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasInvoiceCurrencyCode) {
                missingFields.Add("InvoiceCurrencyCode");
            }

            if (!evidence.HasCurrencyAmount) {
                missingFields.Add("currency-bearing monetary amounts");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-currency-code",
                    "EN 16931 XML currency code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml currency code essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            var invalidCurrencyCodes = new List<string>();
            if (!IsKnownElectronicInvoiceCurrencyCode(evidence.InvoiceCurrencyCode!)) {
                invalidCurrencyCodes.Add(evidence.InvoiceCurrencyCode!.Trim().ToUpperInvariant());
            }

            for (int j = 0; j < evidence.AmountCurrencyCodes.Count; j++) {
                if (!IsKnownElectronicInvoiceCurrencyCode(evidence.AmountCurrencyCodes[j])) {
                    invalidCurrencyCodes.Add(evidence.AmountCurrencyCodes[j].Trim().ToUpperInvariant());
                }
            }

            if (invalidCurrencyCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-currency-code",
                    "EN 16931 XML currency code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml InvoiceCurrencyCode and amount currencyID values to ISO 4217 currency codes before Mustang validation. Found: " + string.Join(", ", invalidCurrencyCodes.Distinct(StringComparer.Ordinal).ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-currency-code",
                "EN 16931 XML currency code",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice invoice and amount currency codes are from the ISO 4217 code list for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML currency codes."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-currency-code",
            "EN 16931 XML currency code",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static bool IsKnownElectronicInvoiceCurrencyCode(string currencyCode) {
        string normalized = currencyCode.Trim().ToUpperInvariant();
        for (int i = 0; i < ElectronicInvoiceCurrencyCodes.Length; i++) {
            if (string.Equals(normalized, ElectronicInvoiceCurrencyCodes[i], StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }
}
