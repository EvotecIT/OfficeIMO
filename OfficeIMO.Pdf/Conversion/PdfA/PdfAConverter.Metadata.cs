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
        "http://www.aiim.org/pdfa/ns/id/part", "http://www.aiim.org/pdfa/ns/id/conformance", "http://www.aiim.org/pdfa/ns/id/amd",
        "http://www.aiim.org/pdfa/ns/id/rev"
    };

    /// <summary>
    /// Writes XMP that identifies the PDF/A part and mirrors the Info dictionary, which PDF/A-2 and PDF/A-3 require to agree.
    /// PDF/A-4 makes XMP authoritative: the Info dictionary is dropped after its entries are carried into XMP, unless the
    /// catalog has PieceInfo, in which case it keeps only ModDate (ISO 19005-4, 6.1.3).
    /// </summary>
    private int? ProcessMetadata(PdfDictionary catalog) {
        int? infoNumber = _document.Security.InfoObjectNumber is int number && _objects.TryGetValue(number, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfDictionary ? number : null;
        PdfDictionary? info = infoNumber is int existingInfo ? (PdfDictionary)_objects[existingInfo].Value : null;
        if (_verifyOnly) {
            VerifyMetadata(catalog, info);
            return infoNumber;
        }
        bool part4 = _part == 4;

        var text = new Dictionary<string, string>(StringComparer.Ordinal);
        var dates = new Dictionary<string, DateTimeOffset>(StringComparer.Ordinal);
        if (info is not null) {
            foreach (string key in TextInfoKeys) {
                if (!info.Items.TryGetValue(key, out PdfObject? value)) continue;
                string? content = Resolve(value) is PdfStringObj stringValue ? stringValue.Value : null;
                if (string.IsNullOrEmpty(content) || !IsXmlText(content!)) {
                    if (part4) continue;
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
                    if (part4) continue;
                    Change(PdfAConversionChangeKind.DocumentInfoUpdated, "document information " + key,
                        "Removed the " + key + " entry because it is not a valid PDF date that XMP metadata can mirror.", info);
                    info.Items.Remove(key);
                    continue;
                }
                dates[key] = parsed.Value;
                string formatted = PdfDateCodec.Format(parsed.Value);
                if (!part4 && !string.Equals(raw, formatted, StringComparison.Ordinal)) {
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
        return part4 ? ReduceInfoForPart4(catalog, info, infoNumber, dates) : infoNumber;
    }

    private int? ReduceInfoForPart4(PdfDictionary catalog, PdfDictionary? info, int? infoNumber, Dictionary<string, DateTimeOffset> dates) {
        bool hasPieceInfo = catalog.Items.ContainsKey("PieceInfo");
        if (!hasPieceInfo) {
            if (info is not null) {
                Change(PdfAConversionChangeKind.DocumentInfoUpdated, "document information",
                    "Removed the document information dictionary, which PDF/A-4 forbids without PieceInfo; its entries are carried in the XMP metadata.", info);
            }
            return null;
        }
        DateTimeOffset modified = dates.TryGetValue("ModDate", out DateTimeOffset modDate) ? modDate
            : dates.TryGetValue("CreationDate", out DateTimeOffset created) ? created : DateTimeOffset.UtcNow;
        var reduced = new PdfDictionary();
        reduced.Items["ModDate"] = new PdfStringObj(PdfDateCodec.Format(modified));
        if (info is not null && info.Items.Count == 1 && info.Items.ContainsKey("ModDate") && dates.ContainsKey("ModDate")) return infoNumber;
        Change(PdfAConversionChangeKind.DocumentInfoUpdated, "document information",
            "Reduced the document information dictionary to ModDate, as PDF/A-4 requires when the catalog has PieceInfo; the other entries are carried in the XMP metadata.",
            (PdfObject?)info ?? catalog);
        if (info is not null && infoNumber is int existing) {
            info.Items.Clear();
            info.Items["ModDate"] = reduced.Items["ModDate"];
            return existing;
        }
        return AddObject(reduced).ObjectNumber;
    }

    private void VerifyMetadata(PdfDictionary catalog, PdfDictionary? info) {
        if (_part == 4 && info is not null &&
            (!catalog.Items.ContainsKey("PieceInfo") || info.Items.Count != 1 || !info.Items.ContainsKey("ModDate"))) {
            AddViolation(PdfAConversionRefusalReason.VerificationFailed, "document information",
                "The converted file keeps document information entries that PDF/A-4 forbids.", info);
        }
        if (Resolve(catalog.Items.TryGetValue("Metadata", out PdfObject? metadata) ? metadata : null) is not PdfStream stream ||
            stream.Dictionary.Items.ContainsKey("Filter")) {
            AddViolation(PdfAConversionRefusalReason.VerificationFailed, "catalog metadata", "The converted file has no uncompressed XMP metadata.", catalog);
            return;
        }
        string xml = Encoding.UTF8.GetString(stream.Data);
        string part = "<pdfaid:part>" + _part.ToString(CultureInfo.InvariantCulture) + "</pdfaid:part>";
        string? conformanceValue = Conformance;
        bool conformanceMatches = conformanceValue is null
            ? !xml.Contains("<pdfaid:conformance>")
            : xml.Contains("<pdfaid:conformance>" + conformanceValue + "</pdfaid:conformance>");
        bool revisionMatches = _part != 4 || xml.Contains("<pdfaid:rev>2020</pdfaid:rev>");
        if (xml.IndexOf(part, StringComparison.Ordinal) < 0 || !conformanceMatches || !revisionMatches) {
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
        xml.Append("<pdfaid:part>").Append(_part.ToString(CultureInfo.InvariantCulture)).Append("</pdfaid:part>\n");
        if (_part == 4) xml.Append("<pdfaid:rev>2020</pdfaid:rev>\n");
        if (Conformance is string conformance) xml.Append("<pdfaid:conformance>").Append(conformance).Append("</pdfaid:conformance>\n");
        xml.Append("</rdf:Description>\n</rdf:RDF>\n</x:xmpmeta>\n");
        // Padding lets later metadata edits happen in place, as the XMP specification recommends.
        for (int line = 0; line < 20; line++) xml.Append(' ', 99).Append('\n');
        xml.Append("<?xpacket end=\"w\"?>");
        return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(xml.ToString());
    }

    /// <summary>The pdfaid:conformance value: B or U for parts 2 and 3; PDF/A-4 states one only for 4f (and 4e).</summary>
    private string? Conformance => _part == 4 ? (_attachmentsProfile ? "F" : null) : _unicode ? "U" : "B";

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
