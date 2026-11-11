using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private const string RdfNamespace = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
    private static readonly string[] TextInfoKeys = { "Title", "Author", "Subject", "Keywords", "Creator", "Producer" };
    private static readonly string[] DateInfoKeys = { "CreationDate", "ModDate" };

    private static readonly HashSet<string> WrittenProperties = new(StringComparer.Ordinal) {
        "http://purl.org/dc/elements/1.1/title", "http://purl.org/dc/elements/1.1/creator", "http://purl.org/dc/elements/1.1/description",
        "http://purl.org/dc/elements/1.1/format", "http://purl.org/dc/elements/1.1/subject",
        "http://ns.adobe.com/pdf/1.3/Keywords", "http://ns.adobe.com/pdf/1.3/Producer", "http://ns.adobe.com/pdf/1.3/Trapped",
        "http://ns.adobe.com/xap/1.0/CreatorTool", "http://ns.adobe.com/xap/1.0/CreateDate", "http://ns.adobe.com/xap/1.0/ModifyDate",
        "http://ns.adobe.com/xap/1.0/MetadataDate",
        "http://www.aiim.org/pdfa/ns/id/part", "http://www.aiim.org/pdfa/ns/id/conformance", "http://www.aiim.org/pdfa/ns/id/amd"
    };

    /// <summary>Writes XMP that identifies the PDF/A part and mirrors the Info dictionary, which PDF/A requires to agree.</summary>
    private int? ProcessMetadata(PdfDictionary catalog) {
        int? infoNumber = _document.Security.InfoObjectNumber is int number && _objects.TryGetValue(number, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfDictionary ? number : null;
        PdfDictionary? info = infoNumber is int existingInfo ? (PdfDictionary)_objects[existingInfo].Value : null;
        if (_verifyOnly) {
            VerifyMetadata(catalog);
            return infoNumber;
        }

        var text = new Dictionary<string, string>(StringComparer.Ordinal);
        var dates = new Dictionary<string, DateTimeOffset>(StringComparer.Ordinal);
        if (info is not null) {
            foreach (string key in TextInfoKeys) {
                if (!info.Items.TryGetValue(key, out PdfObject? value)) continue;
                string? content = Resolve(value) is PdfStringObj stringValue ? stringValue.Value : null;
                if (string.IsNullOrEmpty(content) || !IsXmlText(content!)) {
                    Change(PdfAConversionChangeKind.DocumentInfoUpdated, "document information " + key,
                        "Removed an empty or non-text " + key + " entry that XMP metadata cannot mirror.", info);
                    info.Items.Remove(key);
                    continue;
                }
                text[key] = content!;
            }
            foreach (string key in DateInfoKeys) {
                if (!info.Items.TryGetValue(key, out PdfObject? value)) continue;
                string? raw = Resolve(value) is PdfStringObj stringValue ? stringValue.Value : null;
                DateTimeOffset? parsed = PdfDateCodec.TryParse(raw?.Trim());
                if (parsed is null) {
                    Change(PdfAConversionChangeKind.DocumentInfoUpdated, "document information " + key,
                        "Removed the " + key + " entry because it is not a valid PDF date that XMP metadata can mirror.", info);
                    info.Items.Remove(key);
                    continue;
                }
                dates[key] = parsed.Value;
                string formatted = PdfDateCodec.Format(parsed.Value);
                if (!string.Equals(raw, formatted, StringComparison.Ordinal)) {
                    Change(PdfAConversionChangeKind.DocumentInfoUpdated, "document information " + key,
                        "Wrote " + key + " with an explicit time zone so it matches the XMP date exactly.", info);
                    info.Items[key] = new PdfStringObj(formatted);
                }
            }
        }

        PdfStream? previous = Resolve(catalog.Items.TryGetValue("Metadata", out PdfObject? metadata) ? metadata : null) as PdfStream;
        IReadOnlyList<string> dropped = previous is null ? Array.Empty<string>() : FindUnwrittenProperties(previous);
        byte[] packet = BuildXmp(text, dates);
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("Metadata");
        dictionary.Items["Subtype"] = new PdfName("XML");
        catalog.Items["Metadata"] = AddObject(new PdfStream(dictionary, packet));
        string message = (previous is null ? "Added XMP metadata" : "Replaced the XMP metadata with a packet") +
            " identifying " + PdfAConversionReport.ToProfileName(_options.Profile) + " and mirroring the document information.";
        if (dropped.Count > 0) message += " Properties not carried over: " + string.Join(", ", dropped.Take(8)) + (dropped.Count > 8 ? ", ..." : ".");
        Change(PdfAConversionChangeKind.MetadataReplaced, "catalog metadata", message, catalog,
            dropped.Count > 0 ? OfficeConversionLossKind.Omission : OfficeConversionLossKind.None);
        return infoNumber;
    }

    private void VerifyMetadata(PdfDictionary catalog) {
        if (Resolve(catalog.Items.TryGetValue("Metadata", out PdfObject? metadata) ? metadata : null) is not PdfStream stream ||
            stream.Dictionary.Items.ContainsKey("Filter")) {
            AddViolation(PdfAConversionRefusalReason.VerificationFailed, "catalog metadata", "The converted file has no uncompressed XMP metadata.", catalog);
            return;
        }
        string xml = Encoding.UTF8.GetString(stream.Data);
        string part = "<pdfaid:part>" + _part.ToString(CultureInfo.InvariantCulture) + "</pdfaid:part>";
        string conformance = "<pdfaid:conformance>" + (_unicode ? "U" : "B") + "</pdfaid:conformance>";
        if (xml.IndexOf(part, StringComparison.Ordinal) < 0 || xml.IndexOf(conformance, StringComparison.Ordinal) < 0) {
            AddViolation(PdfAConversionRefusalReason.VerificationFailed, "catalog metadata", "The converted XMP metadata does not identify the requested PDF/A profile.", catalog);
        }
    }

    private byte[] BuildXmp(Dictionary<string, string> text, Dictionary<string, DateTimeOffset> dates) {
        var xml = new StringBuilder();
        xml.Append("<?xpacket begin=\"﻿\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?>\n")
            .Append("<x:xmpmeta xmlns:x=\"adobe:ns:meta/\">\n")
            .Append("<rdf:RDF xmlns:rdf=\"").Append(RdfNamespace).Append("\">\n")
            .Append("<rdf:Description rdf:about=\"\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:pdf=\"http://ns.adobe.com/pdf/1.3/\"")
            .Append(" xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\" xmlns:pdfaid=\"http://www.aiim.org/pdfa/ns/id/\">\n");
        xml.Append("<dc:format>application/pdf</dc:format>\n");
        if (text.TryGetValue("Title", out string? title)) AppendLanguageAlternative(xml, "dc:title", title);
        if (text.TryGetValue("Author", out string? author)) xml.Append("<dc:creator><rdf:Seq><rdf:li>").Append(Escape(author)).Append("</rdf:li></rdf:Seq></dc:creator>\n");
        if (text.TryGetValue("Subject", out string? subject)) AppendLanguageAlternative(xml, "dc:description", subject);
        if (text.TryGetValue("Keywords", out string? keywords)) xml.Append("<pdf:Keywords>").Append(Escape(keywords)).Append("</pdf:Keywords>\n");
        if (text.TryGetValue("Producer", out string? producer)) xml.Append("<pdf:Producer>").Append(Escape(producer)).Append("</pdf:Producer>\n");
        if (text.TryGetValue("Creator", out string? creator)) xml.Append("<xmp:CreatorTool>").Append(Escape(creator)).Append("</xmp:CreatorTool>\n");
        if (dates.TryGetValue("CreationDate", out DateTimeOffset created)) xml.Append("<xmp:CreateDate>").Append(FormatXmpDate(created)).Append("</xmp:CreateDate>\n");
        if (dates.TryGetValue("ModDate", out DateTimeOffset modified)) xml.Append("<xmp:ModifyDate>").Append(FormatXmpDate(modified)).Append("</xmp:ModifyDate>\n");
        xml.Append("<pdfaid:part>").Append(_part.ToString(CultureInfo.InvariantCulture)).Append("</pdfaid:part>\n")
            .Append("<pdfaid:conformance>").Append(_unicode ? "U" : "B").Append("</pdfaid:conformance>\n")
            .Append("</rdf:Description>\n</rdf:RDF>\n</x:xmpmeta>\n");
        // Padding lets later metadata edits happen in place, as the XMP specification recommends.
        for (int line = 0; line < 20; line++) xml.Append(' ', 99).Append('\n');
        xml.Append("<?xpacket end=\"w\"?>");
        return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(xml.ToString());
    }

    private static void AppendLanguageAlternative(StringBuilder xml, string element, string value) =>
        xml.Append('<').Append(element).Append("><rdf:Alt><rdf:li xml:lang=\"x-default\">").Append(Escape(value))
            .Append("</rdf:li></rdf:Alt></").Append(element).Append(">\n");

    private static string FormatXmpDate(DateTimeOffset value) {
        TimeSpan offset = value.Offset;
        string zone = offset == TimeSpan.Zero ? "Z" : (offset < TimeSpan.Zero ? "-" : "+") +
            offset.Duration().Hours.ToString("00", CultureInfo.InvariantCulture) + ":" + offset.Duration().Minutes.ToString("00", CultureInfo.InvariantCulture);
        return value.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture) + zone;
    }

    private static string Escape(string value) => value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    private static bool IsXmlText(string value) {
        foreach (char character in value) {
            if (character < 0x20 && character is not ('\t' or '\n' or '\r')) return false;
            if (character is '￾' or '￿') return false;
        }
        return true;
    }

    private IReadOnlyList<string> FindUnwrittenProperties(PdfStream metadata) {
        if (!StreamDecoder.TryDecode(metadata, _document.ReadOptions.Limits.MaxDecodedStreamBytes, out byte[] data, _objects)) {
            return new[] { "(unreadable packet)" };
        }
        try {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = PdfReadDocument.MaxXmpMetadataBytes };
            using var reader = XmlReader.Create(new MemoryStream(data), settings);
            XDocument document = XDocument.Load(reader);
            XNamespace rdf = RdfNamespace;
            var dropped = new SortedSet<string>(StringComparer.Ordinal);
            foreach (XElement description in document.Descendants(rdf + "Description")) {
                foreach (XAttribute attribute in description.Attributes()) {
                    if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace == rdf || attribute.Name.NamespaceName.Length == 0) continue;
                    AddUnwritten(dropped, attribute.Name);
                }
                foreach (XElement property in description.Elements()) AddUnwritten(dropped, property.Name);
            }
            return dropped.ToList();
        } catch (XmlException) {
            return new[] { "(malformed packet)" };
        }
    }

    private static void AddUnwritten(SortedSet<string> dropped, XName name) {
        if (!WrittenProperties.Contains(name.NamespaceName + name.LocalName)) dropped.Add(name.LocalName);
    }
}
