namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    internal static bool TryReadAllowanceChargeReasons(PdfEmbeddedFile file, out PdfCiiAllowanceChargeReasonEvidence? evidence, out string? diagnostic) {
        Guard.NotNull(file, nameof(file));
        evidence = null;

        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                var missingAllowanceReasons = new List<string>();
                var missingChargeReasons = new List<string>();

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

                    if (string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                        ReadAllowanceChargeReasonEssentials(reader, missingAllowanceReasons, missingChargeReasons);
                    }
                }

                if (!sawRoot) {
                    diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }

                evidence = new PdfCiiAllowanceChargeReasonEvidence(missingAllowanceReasons, missingChargeReasons);
                diagnostic = null;
                return true;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static void ReadAllowanceChargeReasonEssentials(
        System.Xml.XmlReader reader,
        List<string> missingAllowanceReasons,
        List<string> missingChargeReasons) {
        if (reader.IsEmptyElement) {
            return;
        }

        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1 &&
                string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                ReadAllowanceChargeReason(reader, missingAllowanceReasons, missingChargeReasons);
                continue;
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "ApplicableHeaderTradeSettlement", StringComparison.Ordinal)) {
                break;
            }
        }
    }

    private static void ReadAllowanceChargeReason(
        System.Xml.XmlReader reader,
        List<string> missingAllowanceReasons,
        List<string> missingChargeReasons) {
        if (reader.IsEmptyElement) {
            missingAllowanceReasons.Add("document-level allowance");
            return;
        }

        bool? isCharge = null;
        string? actualAmount = null;
        bool hasReason = false;
        int depth = reader.Depth;
        while (reader.Read()) {
            if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                reader.Depth == depth + 1) {
                if (string.Equals(reader.LocalName, "ChargeIndicator", StringComparison.Ordinal)) {
                    ReadChargeIndicator(reader, ref isCharge);
                    continue;
                }

                if (string.Equals(reader.LocalName, "ActualAmount", StringComparison.Ordinal)) {
                    actualAmount = ReadElementText(reader);
                    continue;
                }

                if (string.Equals(reader.LocalName, "Reason", StringComparison.Ordinal) ||
                    string.Equals(reader.LocalName, "ReasonCode", StringComparison.Ordinal)) {
                    hasReason = hasReason || !string.IsNullOrWhiteSpace(ReadElementText(reader));
                    continue;
                }
            }

            if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                reader.Depth == depth &&
                string.Equals(reader.LocalName, "SpecifiedTradeAllowanceCharge", StringComparison.Ordinal)) {
                break;
            }
        }

        if (hasReason) {
            return;
        }

        string marker = string.IsNullOrWhiteSpace(actualAmount)
            ? "without ActualAmount"
            : "ActualAmount " + actualAmount;
        if (isCharge == true) {
            missingChargeReasons.Add("document-level charge " + marker);
        } else {
            missingAllowanceReasons.Add("document-level allowance " + marker);
        }
    }
}
