namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static readonly string[] ElectronicInvoiceCountryCodes = {
        "AD", "AE", "AF", "AG", "AI", "AL", "AM", "AO", "AQ", "AR",
        "AS", "AT", "AU", "AW", "AX", "AZ", "BA", "BB", "BD", "BE",
        "BF", "BG", "BH", "BI", "BJ", "BL", "BM", "BN", "BO", "BQ",
        "BR", "BS", "BT", "BV", "BW", "BY", "BZ", "CA", "CC", "CD",
        "CF", "CG", "CH", "CI", "CK", "CL", "CM", "CN", "CO", "CR",
        "CU", "CV", "CW", "CX", "CY", "CZ", "DE", "DJ", "DK", "DM",
        "DO", "DZ", "EC", "EE", "EG", "EH", "ER", "ES", "ET", "FI",
        "FJ", "FK", "FM", "FO", "FR", "GA", "GB", "GD", "GE", "GF",
        "GG", "GH", "GI", "GL", "GM", "GN", "GP", "GQ", "GR", "GS",
        "GT", "GU", "GW", "GY", "HK", "HM", "HN", "HR", "HT", "HU",
        "ID", "IE", "IL", "IM", "IN", "IO", "IQ", "IR", "IS", "IT",
        "JE", "JM", "JO", "JP", "KE", "KG", "KH", "KI", "KM", "KN",
        "KP", "KR", "KW", "KY", "KZ", "LA", "LB", "LC", "LI", "LK",
        "LR", "LS", "LT", "LU", "LV", "LY", "MA", "MC", "MD", "ME",
        "MF", "MG", "MH", "MK", "ML", "MM", "MN", "MO", "MP", "MQ",
        "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA",
        "NC", "NE", "NF", "NG", "NI", "NL", "NO", "NP", "NR", "NU",
        "NZ", "OM", "PA", "PE", "PF", "PG", "PH", "PK", "PL", "PM",
        "PN", "PR", "PS", "PT", "PW", "PY", "QA", "RE", "RO", "RS",
        "RU", "RW", "SA", "SB", "SC", "SD", "SE", "SG", "SH", "SI",
        "SJ", "SK", "SL", "SM", "SN", "SO", "SR", "SS", "ST", "SV",
        "SX", "SY", "SZ", "TC", "TD", "TF", "TG", "TH", "TJ", "TK",
        "TL", "TM", "TN", "TO", "TR", "TT", "TV", "TW", "TZ", "UA",
        "UG", "UM", "US", "UY", "UZ", "VA", "VC", "VE", "VG", "VI",
        "VN", "VU", "WF", "WS", "XI", "YE", "YT", "ZA", "ZM", "ZW"
    };

    private static PdfComplianceRequirement BuildElectronicInvoiceXmlCountryCodeRequirement(PdfOptions options) {
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = options.EmbeddedFiles;
        var diagnostics = new List<string>();
        for (int i = 0; i < embeddedFiles.Count; i++) {
            PdfEmbeddedFile file = embeddedFiles[i];
            if (!IsFacturXCiiAttachment(file, diagnostics)) {
                continue;
            }

            if (!PdfCiiDocumentHeaderInspector.TryReadCountryCodes(file, out PdfCiiCountryCodeEvidence? evidence, out string? countryDiagnostic)) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-country-code",
                    "EN 16931 XML country code",
                    PdfComplianceRequirementStatus.Missing,
                    countryDiagnostic!);
            }

            var missingFields = new List<string>();
            if (!evidence!.HasSellerCountryId) {
                missingFields.Add("SellerTradeParty PostalTradeAddress CountryID");
            }

            if (!evidence.HasBuyerCountryId) {
                missingFields.Add("BuyerTradeParty PostalTradeAddress CountryID");
            }

            if (missingFields.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-country-code",
                    "EN 16931 XML country code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml party country-code essentials before Mustang validation: " + string.Join(", ", missingFields.ToArray()) + ".");
            }

            var invalidCountryCodes = new List<string>();
            AddInvalidCountryCode(evidence.SellerCountryId!, invalidCountryCodes);
            AddInvalidCountryCode(evidence.BuyerCountryId!, invalidCountryCodes);

            if (invalidCountryCodes.Count > 0) {
                return new PdfComplianceRequirement(
                    "einvoice-xml-country-code",
                    "EN 16931 XML country code",
                    PdfComplianceRequirementStatus.Missing,
                    "Set factur-x.xml party CountryID values to ISO 3166-1 alpha-2 country codes before Mustang validation. Found: " + string.Join(", ", invalidCountryCodes.Distinct(StringComparer.Ordinal).ToArray()) + ".");
            }

            return new PdfComplianceRequirement(
                "einvoice-xml-country-code",
                "EN 16931 XML country code",
                PdfComplianceRequirementStatus.Satisfied,
                "The factur-x.xml CrossIndustryInvoice seller and buyer country identifiers are from the ISO 3166-1 alpha-2 code list for e-invoice readiness.");
        }

        string diagnostic = diagnostics.Count == 0
            ? "Attach a canonical factur-x.xml CrossIndustryInvoice file before checking EN 16931 XML country codes."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "einvoice-xml-country-code",
            "EN 16931 XML country code",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static void AddInvalidCountryCode(string countryCode, List<string> invalidCountryCodes) {
        if (!IsKnownElectronicInvoiceCountryCode(countryCode)) {
            invalidCountryCodes.Add(countryCode.Trim().ToUpperInvariant());
        }
    }

    private static bool IsKnownElectronicInvoiceCountryCode(string countryCode) {
        string normalized = countryCode.Trim().ToUpperInvariant();
        for (int i = 0; i < ElectronicInvoiceCountryCodes.Length; i++) {
            if (string.Equals(normalized, ElectronicInvoiceCountryCodes[i], StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }
}
