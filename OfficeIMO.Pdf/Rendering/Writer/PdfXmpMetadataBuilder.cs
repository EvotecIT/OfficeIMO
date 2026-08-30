namespace OfficeIMO.Pdf;

internal static class PdfXmpMetadataBuilder {
    private static readonly char[] KeywordSeparators = { ',', ';' };

    internal static byte[] Build(string? title, string? author, string? subject, string? keywords, PdfAIdentification? pdfAIdentification = null, PdfUaIdentification? pdfUaIdentification = null, PdfElectronicInvoiceMetadata? electronicInvoiceMetadata = null, PdfXIdentification? pdfXIdentification = null, PdfXProductionMetadata? pdfXProductionMetadata = null, PdfTrappingStatus? trappingStatus = null) {
        var sb = new StringBuilder();
        sb.Append("<?xpacket begin=\"\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?>\n");
        sb.Append("<x:xmpmeta xmlns:x=\"adobe:ns:meta/\">\n");
        sb.Append("<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">\n");
        sb.Append("<rdf:Description rdf:about=\"\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:pdf=\"http://ns.adobe.com/pdf/1.3/\" xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\"");
        if (pdfAIdentification != null) {
            sb.Append(" xmlns:pdfaid=\"http://www.aiim.org/pdfa/ns/id/\"");
        }

        if (pdfUaIdentification != null) {
            sb.Append(" xmlns:pdfuaid=\"")
                .Append(PdfUaIdentification.NamespaceUri)
                .Append('"');
        }

        if (pdfXIdentification != null) {
            sb.Append(" xmlns:pdfxid=\"")
                .Append(PdfXIdentification.NamespaceUri)
                .Append('"');
        }

        if (pdfXProductionMetadata != null) {
            sb.Append(" xmlns:xmpMM=\"http://ns.adobe.com/xap/1.0/mm/\"");
        }

        if (electronicInvoiceMetadata != null) {
            sb.Append(" xmlns:fx=\"")
                .Append(PdfElectronicInvoiceMetadata.FacturXNamespaceUri)
                .Append("\" xmlns:pdfaExtension=\"http://www.aiim.org/pdfa/ns/extension/\" xmlns:pdfaSchema=\"http://www.aiim.org/pdfa/ns/schema#\" xmlns:pdfaProperty=\"http://www.aiim.org/pdfa/ns/property#\"");
        }

        sb.Append(">\n");
        AppendAltText(sb, "dc:title", title);
        AppendSeqText(sb, "dc:creator", author);
        AppendAltText(sb, "dc:description", subject);
        AppendSubjectBag(sb, keywords);
        sb.Append("<dc:format>application/pdf</dc:format>\n");
        sb.Append("<pdf:Producer>OfficeIMO.Pdf</pdf:Producer>\n");
        AppendElement(sb, "pdf:Keywords", keywords);
        AppendPdfAIdentification(sb, pdfAIdentification);
        AppendPdfUaIdentification(sb, pdfUaIdentification);
        AppendPdfXIdentification(sb, pdfXIdentification);
        AppendPdfXProductionMetadata(sb, pdfXProductionMetadata, trappingStatus);
        AppendElectronicInvoiceMetadata(sb, electronicInvoiceMetadata);
        sb.Append("</rdf:Description>\n");
        sb.Append("</rdf:RDF>\n");
        sb.Append("</x:xmpmeta>\n");
        sb.Append("<?xpacket end=\"w\"?>\n");
        return Encoding.UTF8.GetBytes(sb.ToString());
    }

    private static void AppendAltText(StringBuilder sb, string elementName, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        sb.Append('<').Append(elementName).Append("><rdf:Alt><rdf:li xml:lang=\"x-default\">")
            .Append(EscapeXml(value!))
            .Append("</rdf:li></rdf:Alt></")
            .Append(elementName)
            .Append(">\n");
    }

    private static void AppendSeqText(StringBuilder sb, string elementName, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        sb.Append('<').Append(elementName).Append("><rdf:Seq><rdf:li>")
            .Append(EscapeXml(value!))
            .Append("</rdf:li></rdf:Seq></")
            .Append(elementName)
            .Append(">\n");
    }

    private static void AppendSubjectBag(StringBuilder sb, string? keywords) {
        var values = SplitKeywords(keywords);
        if (values.Count == 0) {
            return;
        }

        sb.Append("<dc:subject><rdf:Bag>");
        foreach (string value in values) {
            sb.Append("<rdf:li>")
                .Append(EscapeXml(value))
                .Append("</rdf:li>");
        }

        sb.Append("</rdf:Bag></dc:subject>\n");
    }

    private static void AppendElement(StringBuilder sb, string elementName, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        sb.Append('<')
            .Append(elementName)
            .Append('>')
            .Append(EscapeXml(value!))
            .Append("</")
            .Append(elementName)
            .Append(">\n");
    }

    private static void AppendPdfAIdentification(StringBuilder sb, PdfAIdentification? identification) {
        if (identification == null) {
            return;
        }

        sb.Append("<pdfaid:part>")
            .Append(identification.Part.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append("</pdfaid:part>\n");
        if (!string.IsNullOrEmpty(identification.Conformance)) {
            sb.Append("<pdfaid:conformance>")
                .Append(identification.Conformance)
                .Append("</pdfaid:conformance>\n");
        }
    }

    private static void AppendPdfUaIdentification(StringBuilder sb, PdfUaIdentification? identification) {
        if (identification == null) {
            return;
        }

        sb.Append("<pdfuaid:part>")
            .Append(identification.Part.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append("</pdfuaid:part>\n");
    }

    private static void AppendPdfXIdentification(StringBuilder sb, PdfXIdentification? identification) {
        if (identification == null) {
            return;
        }

        AppendElement(sb, "pdfxid:GTS_PDFXVersion", identification.Version);
        AppendElement(sb, "pdfxid:GTS_PDFXConformance", identification.Conformance);
    }

    private static void AppendPdfXProductionMetadata(
        StringBuilder sb,
        PdfXProductionMetadata? metadata,
        PdfTrappingStatus? trappingStatus) {
        if (metadata == null) {
            return;
        }

        AppendElement(sb, "xmp:CreateDate", FormatXmpDate(metadata.CreationDate));
        AppendElement(sb, "xmp:ModifyDate", FormatXmpDate(metadata.ModificationDate));
        AppendElement(sb, "xmp:MetadataDate", FormatXmpDate(metadata.ModificationDate));
        AppendElement(sb, "xmp:CreatorTool", "OfficeIMO.Pdf");
        AppendElement(sb, "xmpMM:DocumentID", FormatUuid(metadata.DocumentId));
        AppendElement(sb, "xmpMM:InstanceID", FormatUuid(metadata.InstanceId));
        AppendElement(sb, "xmpMM:VersionID", metadata.VersionId);
        AppendElement(sb, "xmpMM:RenditionClass", metadata.RenditionClass);
        if (trappingStatus == PdfTrappingStatus.True || trappingStatus == PdfTrappingStatus.False) {
            AppendElement(sb, "pdf:Trapped", trappingStatus == PdfTrappingStatus.True ? "True" : "False");
        }
    }

    private static string FormatXmpDate(DateTimeOffset value) =>
        value.ToString("yyyy-MM-dd'T'HH:mm:sszzz", System.Globalization.CultureInfo.InvariantCulture);

    private static string FormatUuid(Guid value) => "uuid:" + value.ToString("D");

    private static void AppendElectronicInvoiceMetadata(StringBuilder sb, PdfElectronicInvoiceMetadata? metadata) {
        if (metadata == null) {
            return;
        }

        AppendElement(sb, "fx:DocumentType", metadata.DocumentType);
        AppendElement(sb, "fx:DocumentFileName", metadata.DocumentFileName);
        AppendElement(sb, "fx:Version", metadata.Version);
        AppendElement(sb, "fx:ConformanceLevel", metadata.ConformanceLevel);
        AppendFacturXExtensionSchema(sb);
    }

    private static void AppendFacturXExtensionSchema(StringBuilder sb) {
        sb.Append("<pdfaExtension:schemas><rdf:Bag><rdf:li rdf:parseType=\"Resource\">");
        sb.Append("<pdfaSchema:schema>Factur-X PDF/A Extension Schema</pdfaSchema:schema>");
        sb.Append("<pdfaSchema:namespaceURI>")
            .Append(PdfElectronicInvoiceMetadata.FacturXNamespaceUri)
            .Append("</pdfaSchema:namespaceURI>");
        sb.Append("<pdfaSchema:prefix>fx</pdfaSchema:prefix>");
        sb.Append("<pdfaSchema:property><rdf:Seq>");
        AppendFacturXExtensionProperty(sb, "DocumentType", "Document type");
        AppendFacturXExtensionProperty(sb, "DocumentFileName", "Name of the embedded XML invoice file");
        AppendFacturXExtensionProperty(sb, "Version", "Factur-X/ZUGFeRD schema version");
        AppendFacturXExtensionProperty(sb, "ConformanceLevel", "Factur-X/ZUGFeRD conformance level");
        sb.Append("</rdf:Seq></pdfaSchema:property>");
        sb.Append("</rdf:li></rdf:Bag></pdfaExtension:schemas>\n");
    }

    private static void AppendFacturXExtensionProperty(StringBuilder sb, string name, string description) {
        sb.Append("<rdf:li rdf:parseType=\"Resource\">");
        sb.Append("<pdfaProperty:name>")
            .Append(name)
            .Append("</pdfaProperty:name>");
        sb.Append("<pdfaProperty:valueType>Text</pdfaProperty:valueType>");
        sb.Append("<pdfaProperty:category>external</pdfaProperty:category>");
        sb.Append("<pdfaProperty:description>")
            .Append(EscapeXml(description))
            .Append("</pdfaProperty:description>");
        sb.Append("</rdf:li>");
    }

    private static List<string> SplitKeywords(string? keywords) {
        var values = new List<string>();
        if (string.IsNullOrWhiteSpace(keywords)) {
            return values;
        }

        string[] parts = keywords!.Split(KeywordSeparators, StringSplitOptions.RemoveEmptyEntries);
        foreach (string part in parts) {
            string value = part.Trim();
            if (value.Length > 0) {
                values.Add(value);
            }
        }

        if (values.Count == 0) {
            values.Add(keywords.Trim());
        }

        return values;
    }

    private static string EscapeXml(string value) {
        return value
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");
    }
}
