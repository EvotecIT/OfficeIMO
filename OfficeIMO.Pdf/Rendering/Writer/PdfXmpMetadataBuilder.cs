namespace OfficeIMO.Pdf;

internal static class PdfXmpMetadataBuilder {
    private static readonly char[] KeywordSeparators = { ',', ';' };

    internal static byte[] Build(string? title, string? author, string? subject, string? keywords) {
        var sb = new StringBuilder();
        sb.Append("<?xpacket begin=\"\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?>\n");
        sb.Append("<x:xmpmeta xmlns:x=\"adobe:ns:meta/\">\n");
        sb.Append("<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">\n");
        sb.Append("<rdf:Description rdf:about=\"\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:pdf=\"http://ns.adobe.com/pdf/1.3/\" xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\">\n");
        AppendAltText(sb, "dc:title", title);
        AppendSeqText(sb, "dc:creator", author);
        AppendAltText(sb, "dc:description", subject);
        AppendSubjectBag(sb, keywords);
        sb.Append("<pdf:Producer>OfficeIMO.Pdf</pdf:Producer>\n");
        AppendElement(sb, "pdf:Keywords", keywords);
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
