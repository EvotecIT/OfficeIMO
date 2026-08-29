using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal static class EndNoteXmlCodec {
    private static readonly HashSet<string> KnownRecordElements = new HashSet<string>(new[] {
        "rec-number", "ref-type", "contributors", "titles", "periodical", "pages", "volume", "number", "edition", "publisher", "pub-location", "abstract", "language", "dates", "isbn", "electronic-resource-num", "accession-num", "urls", "keywords", "notes"
    }, StringComparer.OrdinalIgnoreCase);

    internal static IList<BibliographyItem> Parse(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, IList<BibliographyNativeEntry> nativeEntries, CancellationToken cancellationToken) {
        var items = new List<BibliographyItem>();
        var limits = new BibliographyLimitGuard(options);
        try {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = options.MaximumInputCharacters };
            ValidateDepth(source, settings, limits, items, cancellationToken);
            using var textReader = new StringReader(source);
            using XmlReader reader = XmlReader.Create(textReader, settings);
            XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
            XElement? root = document.Root;
            if (root != null) foreach (XElement element in root.Elements().Where(element => element.Name.LocalName != "records")) nativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "element", element.ToString(SaveOptions.DisableFormatting), element.Name.LocalName));
            foreach (XElement record in document.Descendants().Where(element => element.Name.LocalName == "record")) {
                cancellationToken.ThrowIfCancellationRequested();
                limits.AddItem(items, GetOffset(record));
                BibliographyItem item = ParseRecord(record, items, limits, diagnostics);
                items.Add(item);
            }
            if (items.Count == 0) diagnostics.Add(new BibliographyDiagnostic("BIBEND001", BibliographyDiagnosticSeverity.Warning, "EndNote XML contains no record elements."));
        } catch (XmlException exception) {
            diagnostics.Add(new BibliographyDiagnostic("BIBEND002", BibliographyDiagnosticSeverity.Error, exception.Message, line: exception.LineNumber, column: exception.LinePosition));
        }
        return items;
    }

    private static void ValidateDepth(string source, XmlReaderSettings settings, BibliographyLimitGuard limits, IList<BibliographyItem> items, CancellationToken cancellationToken) {
        using var textReader = new StringReader(source);
        using XmlReader reader = XmlReader.Create(textReader, settings);
        while (reader.Read()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element) limits.CheckDepth(items, reader.Depth + 1, 0);
        }
    }

    internal static string Write(BibliographyDocument document, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        var settings = new XmlWriterSettings { Encoding = options.Encoding, Indent = true, IndentChars = "  ", NewLineChars = options.LineEnding, NewLineHandling = NewLineHandling.Replace, OmitXmlDeclaration = false };
        var builder = new StringBuilder();
        using (var textWriter = new EncodingStringWriter(builder, options.Encoding))
        using (XmlWriter writer = XmlWriter.Create(textWriter, settings)) {
            writer.WriteStartDocument(); writer.WriteStartElement("xml");
            foreach (BibliographyNativeEntry entry in document.NativeEntries.Where(entry => entry.Format == BibliographyFormat.EndNoteXml && entry.Kind == "element")) {
                if (TryWriteElement(writer, entry.Value)) report.Add("BIBCONV015", BibliographyDiagnosticSeverity.Information, $"Preserved document-level EndNote XML element '{entry.Name}'.", BibliographyConversionAction.PreservedExtension, field: entry.Name);
                else report.Add("BIBCONV117", BibliographyDiagnosticSeverity.Warning, $"Document-level EndNote XML element '{entry.Name}' is malformed and was omitted.", BibliographyConversionAction.Omitted, field: entry.Name);
            }
            writer.WriteStartElement("records");
            for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
                BibliographyItem item = document.Items[itemIndex];
                cancellationToken.ThrowIfCancellationRequested();
                writer.WriteStartElement("record");
                writer.WriteElementString("rec-number", CodecMappings.OutputKey(item, itemIndex));
                writer.WriteStartElement("ref-type"); writer.WriteAttributeString("name", ToEndNoteType(item.Type)); writer.WriteString(ToEndNoteNumber(item.Type).ToString(CultureInfo.InvariantCulture)); writer.WriteEndElement();
                WriteContributors(writer, item); WriteTitles(writer, item); WriteElement(writer, "pages", item.Pages); WriteElement(writer, "volume", item.Volume); WriteElement(writer, "number", item.Issue);
                WriteElement(writer, "edition", item.Edition); WriteElement(writer, "publisher", item.Publisher); WriteElement(writer, "pub-location", item.PublisherPlace);
                WriteElement(writer, "abstract", item.Abstract); WriteElement(writer, "language", item.Language); WriteDates(writer, item);
                foreach (BibliographyIdentifier identifier in item.Identifiers) WriteIdentifier(writer, identifier);
                if (!string.IsNullOrWhiteSpace(item.Url)) { writer.WriteStartElement("urls"); writer.WriteStartElement("related-urls"); WriteElement(writer, "url", item.Url); writer.WriteEndElement(); writer.WriteEndElement(); }
                if (item.Keywords.Count > 0) { writer.WriteStartElement("keywords"); foreach (string keyword in item.Keywords) WriteElement(writer, "keyword", keyword); writer.WriteEndElement(); }
                if (item.Notes.Count > 0) WriteElement(writer, "notes", string.Join("; ", item.Notes));
                foreach (BibliographyNativeField field in item.NativeFields) {
                    if (field.Format == BibliographyFormat.EndNoteXml && !KnownRecordElements.Contains(field.Name) && TryWriteElement(writer, field.RawValue ?? field.Value)) {
                        report.Add("BIBCONV014", BibliographyDiagnosticSeverity.Information, $"Preserved native EndNote XML element '{field.Name}'.", BibliographyConversionAction.PreservedExtension, item, field.Name);
                    } else if (field.Format != BibliographyFormat.EndNoteXml) {
                        report.Add("BIBCONV115", BibliographyDiagnosticSeverity.Warning, $"Native {field.Format} field '{field.Name}' cannot be represented safely in EndNote XML.", BibliographyConversionAction.Omitted, item, field.Name);
                    } else {
                        report.Add("BIBCONV123", BibliographyDiagnosticSeverity.Warning, $"Native EndNote XML field '{field.Name}' conflicts with a typed element or is malformed.", BibliographyConversionAction.Omitted, item, field.Name);
                    }
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); writer.WriteEndElement(); writer.WriteEndDocument();
        }
        foreach (BibliographyNativeEntry entry in document.NativeEntries.Where(entry => entry.Format != BibliographyFormat.EndNoteXml || entry.Kind != "element")) report.Add("BIBCONV116", BibliographyDiagnosticSeverity.Warning, $"Document-level {entry.Format} entry '{entry.Kind}' cannot be represented in EndNote XML.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
        return builder.ToString();
    }

    private static BibliographyItem ParseRecord(XElement record, IList<BibliographyItem> partial, BibliographyLimitGuard limits, List<BibliographyDiagnostic> diagnostics) {
        foreach (XElement leaf in record.Descendants().Where(static element => !element.HasElements)) limits.AddValue(partial, leaf.Value, GetOffset(leaf));
        string type = Child(record, "ref-type")?.Attribute("name")?.Value ?? Value(record, "ref-type");
        var item = new BibliographyItem { Key = Value(record, "rec-number"), NativeType = type, Type = CodecMappings.ParseType(type) };
        XElement? titles = Child(record, "titles");
        item.Title = Value(titles, "title"); item.ContainerTitle = FirstNonEmpty(Value(titles, "secondary-title"), Value(Child(record, "periodical"), "full-title")); item.CollectionTitle = Value(titles, "tertiary-title");
        item.Pages = Value(record, "pages"); item.Volume = Value(record, "volume"); item.Issue = Value(record, "number"); item.Edition = Value(record, "edition");
        item.Publisher = Value(record, "publisher"); item.PublisherPlace = Value(record, "pub-location"); item.Abstract = Value(record, "abstract"); item.Language = Value(record, "language");
        ParseContributors(item, Child(record, "contributors")); ParseDates(item, Child(record, "dates"));
        foreach (XElement identifier in record.Elements().Where(element => element.Name.LocalName == "isbn")) AddIdentifier(item, CodecMappings.InferSerialScheme(identifier.Value), identifier.Value);
        foreach (XElement identifier in record.Elements().Where(element => element.Name.LocalName == "electronic-resource-num")) AddIdentifier(item, "DOI", identifier.Value);
        foreach (XElement identifier in record.Elements().Where(element => element.Name.LocalName == "accession-num")) ParseAccessionIdentifier(item, identifier.Value);
        XElement? urls = Child(record, "urls"); item.Url = urls?.Descendants().FirstOrDefault(element => element.Name.LocalName == "url")?.Value;
        XElement? keywords = Child(record, "keywords"); if (keywords != null) foreach (XElement keyword in keywords.Elements().Where(element => element.Name.LocalName == "keyword")) item.Keywords.Add(keyword.Value);
        string note = Value(record, "notes"); if (!string.IsNullOrWhiteSpace(note)) item.Notes.Add(note);
        foreach (XElement element in record.Elements()) {
            if (!KnownRecordElements.Contains(element.Name.LocalName)) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, element.Name.LocalName, element.Value, element.ToString(SaveOptions.DisableFormatting)));
            else if (HasUnsupportedNestedContent(element)) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, element.Name.LocalName, element.Value, element.ToString(SaveOptions.DisableFormatting)));
        }
        if (string.IsNullOrWhiteSpace(item.Key)) diagnostics.Add(new BibliographyDiagnostic("BIBEND003", BibliographyDiagnosticSeverity.Warning, "EndNote XML record has no rec-number."));
        return item;
    }

    private static void ParseContributors(BibliographyItem item, XElement? contributors) {
        if (contributors == null) return;
        foreach (XElement group in contributors.Elements()) {
            BibliographyContributorRole role = RoleFromElement(group.Name.LocalName);
            foreach (XElement value in group.Elements()) item.Contributors.Add(new BibliographyContributor(role, CodecMappings.ParseCommaName(value.Value)));
        }
    }

    private static void ParseDates(BibliographyItem item, XElement? dates) {
        if (dates == null) return;
        string year = Value(dates, "year"); string pubDate = dates.Descendants().FirstOrDefault(element => element.Name.LocalName == "date")?.Value ?? string.Empty;
        string issued = string.IsNullOrWhiteSpace(year) ? pubDate : string.IsNullOrWhiteSpace(pubDate) ? year : pubDate.StartsWith(year, StringComparison.Ordinal) ? pubDate : year + " " + pubDate;
        if (!string.IsNullOrWhiteSpace(issued)) item.Dates.Add(CodecMappings.ParseDate(BibliographyDateRole.Issued, issued));
    }

    private static void WriteContributors(XmlWriter writer, BibliographyItem item) {
        if (item.Contributors.Count == 0) return;
        writer.WriteStartElement("contributors");
        foreach (IGrouping<BibliographyContributorRole, BibliographyContributor> group in item.Contributors.GroupBy(static contributor => contributor.Role)) {
            writer.WriteStartElement(ElementFromRole(group.Key)); foreach (BibliographyContributor contributor in group) WriteElement(writer, "author", CodecMappings.FormatName(contributor.Name)); writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static void WriteTitles(XmlWriter writer, BibliographyItem item) {
        if (string.IsNullOrWhiteSpace(item.Title) && string.IsNullOrWhiteSpace(item.ContainerTitle) && string.IsNullOrWhiteSpace(item.CollectionTitle)) return;
        writer.WriteStartElement("titles"); WriteElement(writer, "title", item.Title); WriteElement(writer, "secondary-title", item.ContainerTitle); WriteElement(writer, "tertiary-title", item.CollectionTitle); writer.WriteEndElement();
    }

    private static void WriteDates(XmlWriter writer, BibliographyItem item) {
        BibliographyDate? date = item.GetDate(BibliographyDateRole.Issued); if (date == null) return;
        writer.WriteStartElement("dates"); if (date.Year.HasValue) WriteElement(writer, "year", date.Year.Value.ToString(CultureInfo.InvariantCulture));
        string formatted = CodecMappings.FormatDate(date); if (!string.IsNullOrWhiteSpace(formatted)) { writer.WriteStartElement("pub-dates"); WriteElement(writer, "date", formatted); writer.WriteEndElement(); } writer.WriteEndElement();
    }

    private static void WriteIdentifier(XmlWriter writer, BibliographyIdentifier identifier) {
        if (string.Equals(identifier.Scheme, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase)) WriteElement(writer, "isbn", identifier.Value);
        else if (string.Equals(identifier.Scheme, "DOI", StringComparison.OrdinalIgnoreCase)) WriteElement(writer, "electronic-resource-num", identifier.Value);
        else if (string.Equals(identifier.Scheme, "accession", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "PMID", StringComparison.OrdinalIgnoreCase)) WriteElement(writer, "accession-num", identifier.Value);
    }

    private static bool TryWriteElement(XmlWriter writer, string xml) { try { XElement element = XElement.Parse(xml, LoadOptions.PreserveWhitespace); element.WriteTo(writer); return true; } catch (XmlException) { return false; } }
    private static void WriteElement(XmlWriter writer, string name, string? value) { if (!string.IsNullOrWhiteSpace(value)) writer.WriteElementString(name, SanitizeXml(value!)); }
    private static string SanitizeXml(string value) {
        var builder = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) { builder.Append(value[index]).Append(value[++index]); continue; }
            builder.Append(XmlConvert.IsXmlChar(value[index]) ? value[index] : '\uFFFD');
        }
        return builder.ToString();
    }
    private static bool HasUnsupportedNestedContent(XElement element) {
        if (element.Attributes().Any(attribute => !string.Equals(element.Name.LocalName, "ref-type", StringComparison.OrdinalIgnoreCase) || !string.Equals(attribute.Name.LocalName, "name", StringComparison.OrdinalIgnoreCase))) return true;
        foreach (XElement descendant in element.Descendants()) {
            if (descendant.Attributes().Any() || !IsKnownNestedElement(element.Name.LocalName, descendant.Name.LocalName)) return true;
        }
        return false;
    }
    private static bool IsKnownNestedElement(string container, string name) {
        name = name.ToLowerInvariant();
        switch (container.ToLowerInvariant()) {
            case "titles": return name == "title" || name == "secondary-title" || name == "tertiary-title";
            case "periodical": return name == "full-title";
            case "contributors": return name == "authors" || name == "secondary-authors" || name == "tertiary-authors" || name == "subsidiary-authors" || name == "author";
            case "dates": return name == "year" || name == "pub-dates" || name == "date";
            case "urls": return name == "related-urls" || name == "url";
            case "keywords": return name == "keyword";
            default: return false;
        }
    }
    private static XElement? Child(XElement? parent, string name) => parent?.Elements().FirstOrDefault(element => element.Name.LocalName == name);
    private static string Value(XElement? parent, string name) => Child(parent, name)?.Value ?? string.Empty;
    private static string FirstNonEmpty(params string[] values) => values.FirstOrDefault(static value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
    private static int GetOffset(XElement element) { IXmlLineInfo info = element; return info.HasLineInfo() ? info.LineNumber : 0; }
    private static void AddIdentifier(BibliographyItem item, string scheme, string value) { if (!string.IsNullOrWhiteSpace(value)) item.Identifiers.Add(new BibliographyIdentifier(scheme, value)); }
    private static void ParseAccessionIdentifier(BibliographyItem item, string value) {
        AddIdentifier(item, "accession", value);
    }
    private static BibliographyContributorRole RoleFromElement(string name) { switch (name.ToLowerInvariant()) { case "authors": return BibliographyContributorRole.Author; case "secondary-authors": return BibliographyContributorRole.Editor; case "tertiary-authors": return BibliographyContributorRole.CollectionEditor; case "subsidiary-authors": return BibliographyContributorRole.Translator; default: return BibliographyContributorRole.Other; } }
    private static string ElementFromRole(BibliographyContributorRole role) { switch (role) { case BibliographyContributorRole.Author: return "authors"; case BibliographyContributorRole.Editor: return "secondary-authors"; case BibliographyContributorRole.CollectionEditor: return "tertiary-authors"; case BibliographyContributorRole.Translator: return "subsidiary-authors"; default: return "subsidiary-authors"; } }
    private static string ToEndNoteType(BibliographyItemType type) { switch (type) { case BibliographyItemType.ArticleJournal: return "Journal Article"; case BibliographyItemType.Book: return "Book"; case BibliographyItemType.Chapter: return "Book Section"; case BibliographyItemType.PaperConference: return "Conference Paper"; case BibliographyItemType.Report: return "Report"; case BibliographyItemType.Thesis: return "Thesis"; case BibliographyItemType.WebPage: return "Web Page"; case BibliographyItemType.Patent: return "Patent"; default: return "Generic"; } }
    private static int ToEndNoteNumber(BibliographyItemType type) { switch (type) { case BibliographyItemType.ArticleJournal: return 17; case BibliographyItemType.Book: return 6; case BibliographyItemType.Chapter: return 5; case BibliographyItemType.PaperConference: return 47; case BibliographyItemType.Report: return 27; case BibliographyItemType.Thesis: return 32; case BibliographyItemType.WebPage: return 12; case BibliographyItemType.Patent: return 21; default: return 13; } }

    private sealed class EncodingStringWriter : StringWriter {
        private readonly Encoding _encoding;
        internal EncodingStringWriter(StringBuilder builder, Encoding encoding) : base(builder, CultureInfo.InvariantCulture) => _encoding = encoding;
        public override Encoding Encoding => _encoding;
    }
}
