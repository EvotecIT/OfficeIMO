namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadCurrencyConsistency(PdfEmbeddedFile file, out PdfCiiCurrencyConsistencyEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                string? invoiceCurrencyCode = null;
                var amountCurrencyMarkers = new List<PdfCiiCurrencyAmountMarker>();

                while (reader.Read()) {
                    if (reader.NodeType != System.Xml.XmlNodeType.Element) {
                        continue;
                    }

                    if (!sawRoot) {
                        sawRoot = true;
                        if (!IsCiiRoot(reader)) {
                            diagnostic = "Attach UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                            return false;
                        }
                    }

                    if (string.Equals(reader.LocalName, "InvoiceCurrencyCode", StringComparison.Ordinal)) {
                        string value = ReadElementText(reader).Trim();
                        if (!string.IsNullOrWhiteSpace(value)) {
                            invoiceCurrencyCode = value;
                        }

                        continue;
                    }

                    if (IsCurrencyAmountElement(reader.LocalName)) {
                        amountCurrencyMarkers.Add(new PdfCiiCurrencyAmountMarker(reader.LocalName, reader.GetAttribute("currencyID")));
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                BuildCurrencyDiagnostics(
                    invoiceCurrencyCode,
                    amountCurrencyMarkers,
                    out IReadOnlyList<string> amountFieldsWithoutCurrency,
                    out IReadOnlyList<string> mismatchedAmountCurrencyFields);
                evidence = new PdfCiiCurrencyConsistencyEvidence(
                    invoiceCurrencyCode,
                    GetAmountCurrencyCodes(amountCurrencyMarkers),
                    amountCurrencyMarkers.Count > 0,
                    amountFieldsWithoutCurrency,
                    mismatchedAmountCurrencyFields);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void BuildCurrencyDiagnostics(
        string? invoiceCurrencyCode,
        List<PdfCiiCurrencyAmountMarker> amountCurrencyMarkers,
        out IReadOnlyList<string> amountFieldsWithoutCurrency,
        out IReadOnlyList<string> mismatchedAmountCurrencyFields) {
        var missing = new List<string>();
        var mismatched = new List<string>();
        string? normalizedInvoiceCurrency = NormalizeCurrencyCode(invoiceCurrencyCode);

        for (int i = 0; i < amountCurrencyMarkers.Count; i++) {
            PdfCiiCurrencyAmountMarker marker = amountCurrencyMarkers[i];
            string? normalizedAmountCurrency = NormalizeCurrencyCode(marker.CurrencyId);
            if (normalizedAmountCurrency == null) {
                missing.Add(marker.FieldName);
                continue;
            }

            if (normalizedInvoiceCurrency != null &&
                !string.Equals(normalizedAmountCurrency, normalizedInvoiceCurrency, StringComparison.Ordinal)) {
                mismatched.Add(marker.FieldName + " currencyID " + normalizedAmountCurrency);
            }
        }

        amountFieldsWithoutCurrency = missing.Distinct(StringComparer.Ordinal).ToArray();
        mismatchedAmountCurrencyFields = mismatched.Distinct(StringComparer.Ordinal).ToArray();
    }

    private static string[] GetAmountCurrencyCodes(List<PdfCiiCurrencyAmountMarker> amountCurrencyMarkers) {
        var currencyCodes = new List<string>();
        for (int i = 0; i < amountCurrencyMarkers.Count; i++) {
            string? normalized = NormalizeCurrencyCode(amountCurrencyMarkers[i].CurrencyId);
            if (normalized != null) {
                currencyCodes.Add(normalized);
            }
        }

        return currencyCodes.Distinct(StringComparer.Ordinal).ToArray();
    }

    private static string? NormalizeCurrencyCode(string? currencyCode) {
        if (string.IsNullOrWhiteSpace(currencyCode)) {
            return null;
        }

        return currencyCode!.Trim().ToUpperInvariant();
    }

    private static bool IsCurrencyAmountElement(string localName) =>
        string.Equals(localName, "ChargeAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "AllowanceTotalAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "ChargeTotalAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "LineTotalAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "TaxBasisTotalAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "TaxTotalAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "CalculatedAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "BasisAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "GrandTotalAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "DuePayableAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "PaidAmount", StringComparison.Ordinal) ||
        string.Equals(localName, "RoundingAmount", StringComparison.Ordinal);

    private sealed class PdfCiiCurrencyAmountMarker {
        internal PdfCiiCurrencyAmountMarker(string fieldName, string? currencyId) {
            FieldName = fieldName;
            CurrencyId = currencyId;
        }

        internal string FieldName { get; }

        internal string? CurrencyId { get; }
    }
}
