namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadDateFormats(PdfEmbeddedFile file, out PdfCiiDateFormatEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                bool hasIssueDateTime = false;
                bool issueDateTimeIsParseable = false;
                bool hasPaymentDueDateTime = false;
                bool paymentDueDateTimeIsParseable = false;
                var invalidDateFields = new List<string>();

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

                    if (string.Equals(reader.LocalName, "IssueDateTime", StringComparison.Ordinal)) {
                        hasIssueDateTime = true;
                        if (TryReadCiiDateTime(reader, "ExchangedDocument IssueDateTime", invalidDateFields)) {
                            issueDateTimeIsParseable = true;
                        }

                        continue;
                    }

                    if (string.Equals(reader.LocalName, "DueDateDateTime", StringComparison.Ordinal)) {
                        hasPaymentDueDateTime = true;
                        if (TryReadCiiDateTime(reader, "SpecifiedTradePaymentTerms DueDateDateTime", invalidDateFields)) {
                            paymentDueDateTimeIsParseable = true;
                        }
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiDateFormatEvidence(
                    hasIssueDateTime,
                    issueDateTimeIsParseable,
                    hasPaymentDueDateTime,
                    paymentDueDateTimeIsParseable,
                    invalidDateFields.Distinct(StringComparer.Ordinal).ToArray());
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static bool TryReadCiiDateTime(System.Xml.XmlReader reader, string fieldName, List<string> invalidDateFields) {
        string? format = null;
        string? value = null;

        if (reader.IsEmptyElement) {
            invalidDateFields.Add(fieldName + " DateTimeString");
            return false;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                string.Equals(reader.LocalName, "DateTimeString", StringComparison.Ordinal)) {
                format = reader.GetAttribute("format");
                value = ReadElementText(reader);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                (string.Equals(reader.LocalName, "IssueDateTime", StringComparison.Ordinal) ||
                 string.Equals(reader.LocalName, "DueDateDateTime", StringComparison.Ordinal))) {
                break;
            }
        }

        if (IsParseableCiiDateTime(value, format)) {
            return true;
        }

        invalidDateFields.Add(fieldName + " DateTimeString");
        return false;
    }

    private static bool IsParseableCiiDateTime(string? value, string? format) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string trimmedValue = value!.Trim();
        if (string.IsNullOrWhiteSpace(format)) {
            return false;
        }

        string trimmedFormat = format!.Trim();
        string[] exactFormats = GetCiiDateTimeFormats(trimmedFormat);
        if (exactFormats.Length == 0) {
            return false;
        }

        return System.DateTime.TryParseExact(
            trimmedValue,
            exactFormats,
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None,
            out _);
    }

    private static string[] GetCiiDateTimeFormats(string? format) {
        if (string.Equals(format, "102", StringComparison.Ordinal)) {
            return new[] { "yyyyMMdd" };
        }

        if (string.Equals(format, "203", StringComparison.Ordinal)) {
            return new[] { "yyyyMMddHHmm" };
        }

        if (string.Equals(format, "204", StringComparison.Ordinal)) {
            return new[] { "yyyyMMddHHmmss" };
        }

        return Array.Empty<string>();
    }
}
