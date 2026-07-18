namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static void AddElectronicInvoiceRequirements(List<PdfComplianceRequirement> requirements, PdfOptions options) {
        requirements.Add(BuildElectronicInvoiceXmlAttachmentRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlProfileContextRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlDocumentHeaderRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlDocumentTypeCodeRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlDateFormatRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTradeTransactionRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlPartyIdentificationRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlCountryCodeRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlElectronicAddressRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlPartyTaxRegistrationRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlPartyTaxRegistrationSchemeRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlLineItemRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlUnitCodeRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlLinePricingRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlLineAmountConsistencyRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlLineTaxRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlSettlementSummaryRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlCurrencyConsistencyRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlCurrencyCodeRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTaxBreakdownRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTaxCategoryCodeRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTaxCategoryRateRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTaxCategoryAmountRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTaxExemptionReasonRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTaxPartyIdentifierRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTaxCategoryConsistencyRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlTaxTotalConsistencyRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlPaymentInstructionsRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlPaymentMeansCodeRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlPaymentAccountFormatRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlPaymentTermsRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlAmountConsistencyRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlAllowanceChargeReasonRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmlAttachmentParamsRequirement(options));
        requirements.Add(BuildElectronicInvoiceXmpRequirement(options));
        requirements.Add(new PdfComplianceRequirement(
            "mustang-validation",
            "Mustang validation evidence",
            PdfComplianceRequirementStatus.Unsupported,
            "The optional Mustang test gate exists for e-invoice groundwork fixtures, but profile success has not been enabled for generated output."));
    }

    private static PdfComplianceRequirement BuildElectronicInvoiceXmpRequirement(PdfOptions options) {
        PdfElectronicInvoiceMetadata? metadata = options.ElectronicInvoiceMetadata;
        if (metadata == null) {
            return new PdfComplianceRequirement(
                "einvoice-xmp-extension",
                "Factur-X/ZUGFeRD XMP extension metadata",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.SetElectronicInvoiceMetadata(...) so XMP declares the document type, XML file name, schema version, conformance level, and PDF/A extension schema.");
        }

        var diagnostics = new List<string>();
        if (!string.Equals(metadata.DocumentType, "INVOICE", StringComparison.Ordinal)) {
            diagnostics.Add("set e-invoice XMP document type to INVOICE");
        }

        if (!string.Equals(metadata.DocumentFileName, "factur-x.xml", StringComparison.Ordinal)) {
            diagnostics.Add("set e-invoice XMP document file name to factur-x.xml");
        }

        if (string.IsNullOrWhiteSpace(metadata.Version)) {
            diagnostics.Add("set e-invoice XMP version");
        }

        if (string.IsNullOrWhiteSpace(metadata.ConformanceLevel)) {
            diagnostics.Add("set e-invoice XMP conformance level");
        } else if (!IsKnownElectronicInvoiceConformanceLevel(metadata.ConformanceLevel)) {
            diagnostics.Add("set e-invoice XMP conformance level to MINIMUM, BASIC WL, BASIC, EN 16931, EXTENDED, XRECHNUNG, or EXTENDED-CTC-FR");
        }

        if (diagnostics.Count == 0) {
            return new PdfComplianceRequirement(
                "einvoice-xmp-extension",
                "Factur-X/ZUGFeRD XMP extension metadata",
                PdfComplianceRequirementStatus.Satisfied,
                "Factur-X/ZUGFeRD XMP extension metadata is configured with canonical document and attachment fields.");
        }

        return new PdfComplianceRequirement(
            "einvoice-xmp-extension",
            "Factur-X/ZUGFeRD XMP extension metadata",
            PdfComplianceRequirementStatus.Missing,
            "Update the configured e-invoice XMP metadata to " + string.Join(", ", diagnostics.ToArray()) + ".");
    }

    private static bool IsKnownElectronicInvoiceConformanceLevel(string conformanceLevel) {
        string normalized = conformanceLevel.Trim().ToUpperInvariant().Replace("_", " ");
        return normalized == "MINIMUM" ||
            normalized == "BASIC WL" ||
            normalized == "BASIC" ||
            normalized == "EN16931" ||
            normalized == "EN 16931" ||
            normalized == "EXTENDED" ||
            normalized == "XRECHNUNG" ||
            normalized == "EXTENDED-CTC-FR";
    }

    private static bool IsKnownElectronicInvoiceProfileContext(string contextId) {
        string normalized = contextId.Trim().ToUpperInvariant().Replace("_", "").Replace(" ", "");
        return normalized == "URN:FACTUR-X.EU:1P0:MINIMUM" ||
            normalized == "URN:FACTUR-X.EU:1P0:BASICWL" ||
            normalized == "URN:FACTUR-X.EU:1P0:BASIC" ||
            normalized == "URN:FACTUR-X.EU:1P0:EN16931" ||
            normalized == "URN:FACTUR-X.EU:1P0:EXTENDED" ||
            normalized == "URN:FACTUR-X.EU:1P0:XRECHNUNG" ||
            normalized == "URN:FERD:CROSSINDUSTRYDOCUMENT:INVOICE:1P0:BASIC" ||
            normalized == "URN:FERD:CROSSINDUSTRYDOCUMENT:INVOICE:1P0:COMFORT" ||
            normalized == "URN:FERD:CROSSINDUSTRYDOCUMENT:INVOICE:1P0:EXTENDED" ||
            normalized == "URN:CEN.EU:EN16931:2017" ||
            normalized.StartsWith("URN:CEN.EU:EN16931:2017#COMPLIANT#URN:XOEV-DE:KOSIT:STANDARD:XRECHNUNG", StringComparison.Ordinal) ||
            normalized.StartsWith("URN:CEN.EU:EN16931:2017#COMPLIANT#URN:XEINKAUF.DE:KOSIT:XRECHNUNG", StringComparison.Ordinal);
    }

    private static bool IsFacturXCiiAttachment(PdfEmbeddedFile file, List<string> diagnostics) {
        if (!string.Equals(file.FileName, "factur-x.xml", StringComparison.Ordinal)) {
            if (file.FileName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) {
                diagnostics.Add("Use the canonical factur-x.xml attachment filename for Factur-X/ZUGFeRD 2.x readiness.");
            }

            return false;
        }

        if (file.Relationship == PdfAssociatedFileRelationship.Unspecified) {
            diagnostics.Add("Set an explicit associated-file relationship for factur-x.xml.");
            return false;
        }

        if (!IsEInvoiceXmlRelationship(file.Relationship)) {
            diagnostics.Add("Set factur-x.xml AFRelationship to Alternative, Data, or Source.");
            return false;
        }

        if (!IsXmlMimeType(file.MimeType)) {
            diagnostics.Add("Set factur-x.xml MIME type to application/xml or text/xml.");
            return false;
        }

        if (!TryReadCiiXmlRoot(file, out string? rootDiagnostic)) {
            diagnostics.Add(rootDiagnostic!);
            return false;
        }

        return true;
    }

    private static bool IsXmlMimeType(string? mimeType) {
        return string.Equals(mimeType, "application/xml", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(mimeType, "text/xml", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsEInvoiceXmlRelationship(PdfAssociatedFileRelationship relationship) {
        return relationship == PdfAssociatedFileRelationship.Alternative ||
            relationship == PdfAssociatedFileRelationship.Data ||
            relationship == PdfAssociatedFileRelationship.Source;
    }

    private static bool TryReadCiiXmlRoot(PdfEmbeddedFile file, out string? diagnostic) {
        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                while (reader.Read()) {
                    if (reader.NodeType != System.Xml.XmlNodeType.Element) {
                        continue;
                    }

                    if (string.Equals(reader.LocalName, "CrossIndustryInvoice", StringComparison.Ordinal) &&
                        reader.NamespaceURI.StartsWith("urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:", StringComparison.Ordinal)) {
                        diagnostic = null;
                        return true;
                    }

                    diagnostic = "Attach UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                    return false;
                }
            }

            diagnostic = "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
            return false;
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

    private static bool TryReadCiiProfileContext(PdfEmbeddedFile file, out string? contextId, out string? diagnostic) {
        contextId = null;
        try {
            using (var stream = new MemoryStream(file.DataSnapshot))
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            })) {
                bool sawRoot = false;
                int exchangedDocumentContextDepth = -1;
                int guidelineDepth = -1;

                while (reader.Read()) {
                    if (reader.NodeType == System.Xml.XmlNodeType.Element) {
                        if (!sawRoot) {
                            sawRoot = true;
                            if (!string.Equals(reader.LocalName, "CrossIndustryInvoice", StringComparison.Ordinal) ||
                                !reader.NamespaceURI.StartsWith("urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:", StringComparison.Ordinal)) {
                                diagnostic = "Attach UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                                return false;
                            }
                        }

                        if (string.Equals(reader.LocalName, "ExchangedDocumentContext", StringComparison.Ordinal)) {
                            exchangedDocumentContextDepth = reader.Depth;
                            continue;
                        }

                        if (exchangedDocumentContextDepth >= 0 &&
                            string.Equals(reader.LocalName, "GuidelineSpecifiedDocumentContextParameter", StringComparison.Ordinal)) {
                            guidelineDepth = reader.Depth;
                            continue;
                        }

                        if (guidelineDepth >= 0 &&
                            string.Equals(reader.LocalName, "ID", StringComparison.Ordinal)) {
                            string value = reader.ReadElementContentAsString().Trim();
                            if (!string.IsNullOrWhiteSpace(value)) {
                                contextId = value;
                                diagnostic = null;
                                return true;
                            }

                            diagnostic = "Set factur-x.xml ExchangedDocumentContext GuidelineSpecifiedDocumentContextParameter ID to a non-empty profile identifier.";
                            return false;
                        }
                    } else if (reader.NodeType == System.Xml.XmlNodeType.EndElement) {
                        if (reader.Depth == guidelineDepth &&
                            string.Equals(reader.LocalName, "GuidelineSpecifiedDocumentContextParameter", StringComparison.Ordinal)) {
                            guidelineDepth = -1;
                        }

                        if (reader.Depth == exchangedDocumentContextDepth &&
                            string.Equals(reader.LocalName, "ExchangedDocumentContext", StringComparison.Ordinal)) {
                            exchangedDocumentContextDepth = -1;
                        }
                    }
                }

                diagnostic = sawRoot
                    ? "Set factur-x.xml ExchangedDocumentContext GuidelineSpecifiedDocumentContextParameter ID to a recognized Factur-X/ZUGFeRD EN 16931 profile identifier."
                    : "Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.";
                return false;
            }
        } catch (System.Xml.XmlException ex) {
            diagnostic = "Attach parseable XML in factur-x.xml: " + ex.Message;
            return false;
        }
    }

}
