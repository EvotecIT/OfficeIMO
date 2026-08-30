using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal static class EndNoteXmlCodec {
    private const string AttributesEntryKind = "attributes";
    private const string RecordsElementEntryKind = "records-element";
    private const string RecordAttributesFieldName = "@record-attributes";
    private static readonly HashSet<string> KnownRecordElements = new HashSet<string>(new[] {
        "rec-number", "ref-type", "contributors", "titles", "periodical", "pages", "volume", "number", "edition", "publisher", "pub-location", "abstract", "language", "dates", "isbn", "electronic-resource-num", "accession-num", "urls", "keywords", "notes"
    }, StringComparer.OrdinalIgnoreCase);

    internal static IList<BibliographyItem> Parse(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, IList<BibliographyNativeEntry> nativeEntries, out bool recordsRoot, CancellationToken cancellationToken) {
        var items = new List<BibliographyItem>();
        recordsRoot = false;
        var limits = new BibliographyLimitGuard(options);
        var materializationLimits = new BibliographyLimitGuard(options);
        var diagnosticGuard = new BibliographyDiagnosticGuard(options, diagnostics, items);
        try {
            int sourceOffset = source.Length > 0 && source[0] == '\uFEFF' ? 1 : 0;
            string xmlSource = sourceOffset == 0 ? source : source.Substring(sourceOffset);
            var offsets = new EndNoteSourceOffsetMap(xmlSource, sourceOffset, cancellationToken);
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = options.MaximumInputCharacters };
            using var textReader = new StringReader(xmlSource);
            using XmlReader innerReader = XmlReader.Create(textReader, settings);
            using var reader = new EndNoteBoundedXmlReader(innerReader, limits, materializationLimits, items, offsets, cancellationToken);
            XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
            XElement? root = document.Root;
            if (root != null) foreach (XElement element in root.DescendantsAndSelf()) element.AddAnnotation(new EndNoteSourceOffset(offsets.GetOffset(element)));
            bool rootIsRecords = root != null && string.Equals(root.Name.LocalName, "records", StringComparison.OrdinalIgnoreCase);
            recordsRoot = rootIsRecords;
            if (root != null) {
                CaptureAttributes(root, nativeEntries, items, limits);
            }
            if (root != null && !rootIsRecords) foreach (XElement element in root.Elements().Where(element => !HasName(element, root.Name.Namespace, "records"))) {
                ValidateAggregateValueLengths(element, items, limits, true);
                limits.AddValue(items, null, GetOffset(element));
                nativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "element", SerializeBoundedElement(element, items, limits), element.Name.LocalName));
            }
            IEnumerable<XElement> recordContainers = root == null ? Enumerable.Empty<XElement>() : rootIsRecords ? new[] { root } : root.Elements().Where(element => HasName(element, root.Name.Namespace, "records"));
            XElement[] containers = recordContainers.ToArray();
            foreach (XElement container in containers) {
                if (!ReferenceEquals(container, root)) CaptureAttributes(container, nativeEntries, items, limits);
                foreach (XElement element in container.Elements().Where(child => !HasName(child, container.Name.Namespace, "record"))) {
                    ValidateAggregateValueLengths(element, items, limits, true);
                    limits.AddValue(items, null, GetOffset(element));
                    nativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, RecordsElementEntryKind, SerializeBoundedElement(element, items, limits), element.Name.LocalName));
                }
            }
            IEnumerable<XElement> records = containers.SelectMany(element => element.Elements().Where(child => HasName(child, element.Name.Namespace, "record")));
            foreach (XElement record in records) {
                cancellationToken.ThrowIfCancellationRequested();
                BibliographyItem item = ParseRecord(record, items, limits, diagnosticGuard);
                items.Add(item);
            }
            if (items.Count == 0) diagnosticGuard.Add(new BibliographyDiagnostic("BIBEND001", BibliographyDiagnosticSeverity.Warning, "EndNote XML contains no record elements."));
        } catch (XmlException exception) {
            diagnosticGuard.Add(new BibliographyDiagnostic("BIBEND002", BibliographyDiagnosticSeverity.Error, exception.Message, line: exception.LineNumber, column: exception.LinePosition));
        }
        return items;
    }

    internal static string Write(BibliographyDocument document, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        var settings = new XmlWriterSettings { Encoding = options.Encoding, Indent = true, IndentChars = "  ", NewLineChars = options.LineEnding, NewLineHandling = NewLineHandling.Replace, OmitXmlDeclaration = false };
        var builder = new StringBuilder();
        string[] outputKeys = CodecMappings.OutputKeys(document.Items, BibliographyFormat.EndNoteXml, cancellationToken);
        using (var textWriter = new EncodingStringWriter(builder, options.Encoding))
        using (XmlWriter writer = XmlWriter.Create(textWriter, settings)) {
            bool recordsRoot = document.EndNoteRecordsRoot;
            string rootElementName = recordsRoot ? "records" : "xml";
            string outputNamespace = GetDocumentElementNamespace(document, rootElementName, string.Empty, cancellationToken);
            writer.WriteStartDocument();
            if (!recordsRoot) {
                writer.WriteStartElement(null, "xml", outputNamespace);
                WriteDocumentAttributes(writer, document, "xml", report, cancellationToken);
                foreach (BibliographyNativeEntry entry in Cancellable(document.NativeEntries, cancellationToken).Where(entry => entry.Format == BibliographyFormat.EndNoteXml && entry.Kind == "element")) {
                    if (TryWriteRootElement(writer, entry.Value, outputNamespace)) report.Add("BIBCONV015", BibliographyDiagnosticSeverity.Information, $"Preserved document-level EndNote XML element '{entry.Name}'.", BibliographyConversionAction.PreservedExtension, field: entry.Name);
                    else report.Add("BIBCONV117", BibliographyDiagnosticSeverity.Warning, $"Document-level EndNote XML element '{entry.Name}' is malformed or reserved and was omitted.", BibliographyConversionAction.Omitted, field: entry.Name);
                }
            }
            string recordsNamespace = GetDocumentElementNamespace(document, "records", outputNamespace, cancellationToken);
            writer.WriteStartElement(null, "records", recordsNamespace);
            WriteDocumentAttributes(writer, document, "records", report, cancellationToken);
            foreach (BibliographyNativeEntry entry in Cancellable(document.NativeEntries, cancellationToken).Where(entry => entry.Format == BibliographyFormat.EndNoteXml && entry.Kind == RecordsElementEntryKind)) {
                if (TryWriteRecordsElement(writer, entry.Value, recordsNamespace)) report.Add("BIBCONV015", BibliographyDiagnosticSeverity.Information, $"Preserved EndNote XML records-container element '{entry.Name}'.", BibliographyConversionAction.PreservedExtension, field: entry.Name);
                else report.Add("BIBCONV117", BibliographyDiagnosticSeverity.Warning, $"EndNote XML records-container element '{entry.Name}' is malformed, reserved, or otherwise unsafe and was omitted.", BibliographyConversionAction.Omitted, field: entry.Name);
            }
            for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
                BibliographyItem item = document.Items[itemIndex];
                cancellationToken.ThrowIfCancellationRequested();
                string recordNamespace = GetRecordNamespace(item, recordsNamespace, cancellationToken);
                writer.WriteStartElement(null, "record", recordNamespace);
                WriteRecordAttributes(writer, item, report, cancellationToken);
                WriteElement(writer, "rec-number", outputKeys[itemIndex], recordNamespace);
                writer.WriteStartElement(null, "ref-type", recordNamespace); writer.WriteAttributeString("name", OutputType(document.SourceFormat, item)); writer.WriteString(ToEndNoteNumber(item.Type).ToString(CultureInfo.InvariantCulture)); writer.WriteEndElement();
                WriteContributors(writer, item, recordNamespace, cancellationToken); WriteTitles(writer, item, recordNamespace); WritePeriodical(writer, item, recordNamespace); WriteElement(writer, "pages", item.Pages, recordNamespace); WriteElement(writer, "volume", item.Volume, recordNamespace); WriteElement(writer, "number", item.Issue, recordNamespace);
                WriteElement(writer, "edition", item.Edition, recordNamespace); WriteElement(writer, "publisher", item.Publisher, recordNamespace); WriteElement(writer, "pub-location", item.PublisherPlace, recordNamespace);
                WriteElement(writer, "abstract", item.Abstract, recordNamespace); WriteElement(writer, "language", item.Language, recordNamespace); WriteDates(writer, item, report, recordNamespace, cancellationToken);
                foreach (BibliographyIdentifier identifier in Cancellable(item.Identifiers, cancellationToken)) WriteIdentifier(writer, identifier, recordNamespace);
                WriteUrls(writer, item, report, recordNamespace, cancellationToken);
                if (item.Keywords.Count > 0) { writer.WriteStartElement(null, "keywords", recordNamespace); foreach (string keyword in Cancellable(item.Keywords, cancellationToken)) WriteElement(writer, "keyword", keyword, recordNamespace); writer.WriteEndElement(); }
                if (item.Notes.Count > 0) WriteElement(writer, "notes", string.Join("; ", Cancellable(item.Notes, cancellationToken)), recordNamespace);
                foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken)) {
                    if (field.Format == BibliographyFormat.EndNoteXml && string.Equals(field.Name, RecordAttributesFieldName, StringComparison.Ordinal)) {
                        continue;
                    } else if (field.Format == BibliographyFormat.EndNoteXml && IsAdditionalUrlField(field, recordNamespace)) {
                        continue;
                    } else if (field.Format == BibliographyFormat.EndNoteXml && !ConflictsWithTypedRecordElement(item, field, recordNamespace, cancellationToken) && TryWriteNativeField(writer, field, recordNamespace)) {
                        report.Add("BIBCONV014", BibliographyDiagnosticSeverity.Information, $"Preserved native EndNote XML element '{field.Name}'.", BibliographyConversionAction.PreservedExtension, item, field.Name);
                    } else if (field.Format != BibliographyFormat.EndNoteXml) {
                        report.Add("BIBCONV115", BibliographyDiagnosticSeverity.Warning, $"Native {field.Format} field '{field.Name}' cannot be represented safely in EndNote XML.", BibliographyConversionAction.Omitted, item, field.Name);
                    } else {
                        report.Add("BIBCONV123", BibliographyDiagnosticSeverity.Warning, $"Native EndNote XML field '{field.Name}' conflicts with a typed element or is malformed.", BibliographyConversionAction.Omitted, item, field.Name);
                    }
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            if (!recordsRoot) writer.WriteEndElement();
            writer.WriteEndDocument();
        }
        foreach (BibliographyNativeEntry entry in Cancellable(document.NativeEntries, cancellationToken).Where(entry => !IsConsumedEndNoteEntry(entry))) report.Add("BIBCONV116", BibliographyDiagnosticSeverity.Warning, $"Document-level {entry.Format} entry '{entry.Kind}' cannot be represented in EndNote XML.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
        return builder.ToString();
    }

    private static BibliographyItem ParseRecord(XElement record, IList<BibliographyItem> partial, BibliographyLimitGuard limits, BibliographyDiagnosticGuard diagnostics) {
        ValidateAggregateValueLengths(record, partial, limits, false);
        foreach (XElement leaf in record.Descendants().Where(static element => !element.HasElements)) limits.AddValue(partial, leaf.Value, GetOffset(leaf));
        XElement? refType = Child(record, "ref-type");
        string type = refType?.Attribute("name")?.Value ?? refType?.Value ?? string.Empty;
        var item = new BibliographyItem { Key = Value(record, "rec-number"), NativeType = type, Type = CodecMappings.ParseType(type) };
        if (refType?.Attribute("name") != null && item.Type != BibliographyItemType.Unknown &&
            int.TryParse(refType.Value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int nativeTypeNumber) &&
            nativeTypeNumber != ToEndNoteNumber(item.Type))
            diagnostics.Add(new BibliographyDiagnostic("BIBEND004", BibliographyDiagnosticSeverity.Warning, $"EndNote XML ref-type name '{type}' conflicts with numeric code '{nativeTypeNumber}'.", GetOffset(refType), itemKey: item.Key, field: "ref-type"));
        if (HasElementMetadata(record)) {
            string recordMetadata = SerializeAttributes(record);
            limits.AddValue(partial, recordMetadata, GetOffset(record));
            item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, RecordAttributesFieldName, recordMetadata));
        }
        XElement? titles = Child(record, "titles"); XElement? periodical = Child(record, "periodical");
        string? secondaryTitle = OptionalValue(titles, "secondary-title"); string? periodicalTitle = OptionalValue(periodical, "full-title");
        item.Title = OptionalValue(titles, "title"); item.ContainerTitle = FirstNonEmpty(secondaryTitle, periodicalTitle); item.CollectionTitle = OptionalValue(titles, "tertiary-title");
        if (!string.IsNullOrWhiteSpace(secondaryTitle)) item.EndNoteFieldNames["container-title"] = "secondary-title";
        else if (!string.IsNullOrWhiteSpace(periodicalTitle)) item.EndNoteFieldNames["container-title"] = "periodical";
        bool retainedAdditionalPeriodical = periodical != null && !string.IsNullOrWhiteSpace(secondaryTitle) && !string.IsNullOrWhiteSpace(periodicalTitle) && !string.Equals(secondaryTitle, periodicalTitle, StringComparison.Ordinal);
        if (retainedAdditionalPeriodical)
            item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "periodical", periodical!.Value, SerializeBoundedElement(periodical, partial, limits)));
        item.Pages = OptionalValue(record, "pages"); item.Volume = OptionalValue(record, "volume"); item.Issue = OptionalValue(record, "number"); item.Edition = OptionalValue(record, "edition");
        item.Publisher = OptionalValue(record, "publisher"); item.PublisherPlace = OptionalValue(record, "pub-location"); item.Abstract = OptionalValue(record, "abstract"); item.Language = OptionalValue(record, "language");
        ParseContributors(item, Child(record, "contributors")); ParseDates(item, Child(record, "dates"));
        foreach (XElement identifier in record.Elements().Where(element => HasNameInNamespace(element, record.Name.Namespace))) {
            if (string.Equals(identifier.Name.LocalName, "isbn", StringComparison.OrdinalIgnoreCase)) {
                string? declaredScheme = identifier.Attribute("type")?.Value;
                string scheme = string.Equals(declaredScheme, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(declaredScheme, "ISSN", StringComparison.OrdinalIgnoreCase) ? declaredScheme! : CodecMappings.InferSerialScheme(identifier.Value);
                AddIdentifier(item, scheme, identifier.Value);
            } else if (string.Equals(identifier.Name.LocalName, "electronic-resource-num", StringComparison.OrdinalIgnoreCase)) AddIdentifier(item, "DOI", identifier.Value);
            else if (string.Equals(identifier.Name.LocalName, "accession-num", StringComparison.OrdinalIgnoreCase)) ParseAccessionIdentifier(item, identifier.Value);
            if (IsRepeatableRecordElement(identifier.Name.LocalName) && string.IsNullOrWhiteSpace(identifier.Value))
                item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, identifier.Name.LocalName, identifier.Value, SerializeBoundedElement(identifier, partial, limits)));
        }
        XElement? urls = Child(record, "urls"); XElement? relatedUrlsContainer = Child(urls, "related-urls"); XElement[] relatedUrls = relatedUrlsContainer?.Elements().Where(element => HasName(element, record.Name.Namespace, "url")).ToArray() ?? Array.Empty<XElement>();
        string? primaryUrl = relatedUrls.FirstOrDefault()?.Value;
        item.Url = primaryUrl != null && (primaryUrl.Length > 0 || relatedUrls.Length == 1) ? primaryUrl : null;
        foreach (XElement relatedUrl in relatedUrls.Skip(1)) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "url", relatedUrl.Value, SerializeBoundedElement(relatedUrl, partial, limits)));
        XElement? keywords = Child(record, "keywords"); if (keywords != null) foreach (XElement keyword in keywords.Elements().Where(element => HasName(element, keywords.Name.Namespace, "keyword"))) item.Keywords.Add(keyword.Value);
        XElement? note = Child(record, "notes"); if (note != null) item.Notes.Add(note.Value);
        foreach (XElement element in record.Elements()) {
            bool knownRecordElement = HasNameInNamespace(element, record.Name.Namespace) && KnownRecordElements.Contains(element.Name.LocalName);
            bool repeatedSingleValue = knownRecordElement && !IsRepeatableRecordElement(element.Name.LocalName) && element.ElementsBeforeSelf().Any(previous => HasName(previous, record.Name.Namespace, element.Name.LocalName));
            if (!knownRecordElement) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, element.Name.LocalName, element.Value, SerializeBoundedElement(element, partial, limits)));
            else if ((IsEmptyKnownRecordContainer(element) || repeatedSingleValue || HasUnsupportedNestedContent(element) || HasDuplicateKnownNestedContent(element)) && (!ReferenceEquals(element, periodical) || !retainedAdditionalPeriodical)) item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.EndNoteXml, element.Name.LocalName, element.Value, SerializeBoundedElement(element, partial, limits)));
        }
        if (string.IsNullOrWhiteSpace(item.Key)) diagnostics.Add(new BibliographyDiagnostic("BIBEND003", BibliographyDiagnosticSeverity.Warning, "EndNote XML record has no rec-number."));
        return item;
    }

    private static void ParseContributors(BibliographyItem item, XElement? contributors) {
        if (contributors == null) return;
        foreach (XElement group in contributors.Elements().Where(element => HasNameInNamespace(element, contributors.Name.Namespace) && IsContributorRoleElement(element.Name.LocalName))) {
            BibliographyContributorRole role = RoleFromElement(group.Name.LocalName);
            foreach (XElement value in group.Elements().Where(element => HasName(element, group.Name.Namespace, "author"))) item.Contributors.Add(new BibliographyContributor(role, CodecMappings.ParseCommaName(value.Value)));
        }
    }

    private static void ParseDates(BibliographyItem item, XElement? dates) {
        if (dates == null) return;
        XElement? yearElement = Child(dates, "year");
        XElement? dateElement = Child(Child(dates, "pub-dates"), "date");
        if (yearElement == null && dateElement == null) return;
        string year = yearElement?.Value ?? string.Empty; string pubDate = dateElement?.Value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(year) && string.IsNullOrWhiteSpace(pubDate)) {
            XElement retained = dateElement ?? yearElement!;
            item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = retained.Value });
            return;
        }
        BibliographyDate parsedYear = CodecMappings.ParseDate(BibliographyDateRole.Issued, year);
        BibliographyDate parsedPublication = CodecMappings.ParseDate(BibliographyDateRole.Issued, pubDate);
        if (yearElement == null) item.Dates.Add(parsedPublication);
        else if (dateElement == null) item.Dates.Add(parsedYear);
        else if (!parsedYear.Year.HasValue) {
            parsedPublication.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.EndNoteXml, "year", year, yearElement.ToString(SaveOptions.DisableFormatting)));
            item.Dates.Add(parsedPublication);
        }
        else if (parsedYear.Year.HasValue && parsedPublication.Year == parsedYear.Year) item.Dates.Add(parsedPublication);
        else {
            BibliographyDate combined = CodecMappings.ParseDate(BibliographyDateRole.Issued, pubDate + " " + year);
            if (parsedYear.Year.HasValue && combined.Year == parsedYear.Year) item.Dates.Add(combined);
            else { parsedYear.Literal = pubDate; item.Dates.Add(parsedYear); }
        }
    }

    private static void WriteContributors(XmlWriter writer, BibliographyItem item, string xmlNamespace, CancellationToken cancellationToken) {
        if (item.Contributors.Count == 0) return;
        writer.WriteStartElement(null, "contributors", xmlNamespace);
        foreach (IGrouping<BibliographyContributorRole, BibliographyContributor> group in Cancellable(item.Contributors, cancellationToken).GroupBy(static contributor => contributor.Role)) {
            writer.WriteStartElement(null, ElementFromRole(group.Key), xmlNamespace); foreach (BibliographyContributor contributor in Cancellable(group, cancellationToken)) WriteElement(writer, "author", CodecMappings.FormatName(contributor.Name), xmlNamespace); writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static void WriteTitles(XmlWriter writer, BibliographyItem item, string xmlNamespace) {
        string? secondaryTitle = item.EndNoteFieldNames.TryGetValue("container-title", out string? source) && string.Equals(source, "periodical", StringComparison.OrdinalIgnoreCase) ? null : item.ContainerTitle;
        if (item.Title == null && secondaryTitle == null && item.CollectionTitle == null) return;
        writer.WriteStartElement(null, "titles", xmlNamespace); WriteElement(writer, "title", item.Title, xmlNamespace); WriteElement(writer, "secondary-title", secondaryTitle, xmlNamespace); WriteElement(writer, "tertiary-title", item.CollectionTitle, xmlNamespace); writer.WriteEndElement();
    }

    private static void WritePeriodical(XmlWriter writer, BibliographyItem item, string xmlNamespace) {
        if (!item.EndNoteFieldNames.TryGetValue("container-title", out string? source) || !string.Equals(source, "periodical", StringComparison.OrdinalIgnoreCase) || item.ContainerTitle == null) return;
        writer.WriteStartElement(null, "periodical", xmlNamespace); WriteElement(writer, "full-title", item.ContainerTitle, xmlNamespace); writer.WriteEndElement();
    }

    private static void WriteDates(XmlWriter writer, BibliographyItem item, BibliographyConversionReport report, string xmlNamespace, CancellationToken cancellationToken) {
        BibliographyDate? date = Cancellable(item.Dates, cancellationToken).FirstOrDefault(static candidate => candidate.Role == BibliographyDateRole.Issued); if (date == null) return;
        writer.WriteStartElement(null, "dates", xmlNamespace);
        BibliographyNativeField? nativeYear = Cancellable(date.NativeFields, cancellationToken).FirstOrDefault(field => CanPreserveNativeDateField(date, field, cancellationToken));
        if (nativeYear != null) {
            if (TryWriteNativeField(writer, nativeYear, xmlNamespace)) report.Add("BIBCONV014", BibliographyDiagnosticSeverity.Information, "Preserved a distinct EndNote XML year component.", BibliographyConversionAction.PreservedExtension, item, "dates.year");
            else report.Add("BIBCONV123", BibliographyDiagnosticSeverity.Warning, "A distinct EndNote XML year component is malformed and was omitted.", BibliographyConversionAction.Omitted, item, "dates.year");
        } else if (date.Year.HasValue) WriteElement(writer, "year", date.Year.Value.ToString(CultureInfo.InvariantCulture), xmlNamespace);
        string formatted = CodecMappings.FormatDate(date);
        writer.WriteStartElement(null, "pub-dates", xmlNamespace); WriteElement(writer, "date", formatted, xmlNamespace); writer.WriteEndElement();
        writer.WriteEndElement();
    }

    internal static bool CanPreserveNativeDateField(BibliographyDate date, BibliographyNativeField field, CancellationToken cancellationToken = default) {
        if (field.Format != BibliographyFormat.EndNoteXml || !string.Equals(field.Name, "year", StringComparison.OrdinalIgnoreCase)) return false;
        BibliographyDate parsed = CodecMappings.ParseDate(BibliographyDateRole.Issued, field.Value);
        return !parsed.Year.HasValue && !Cancellable(date.NativeFields.TakeWhile(candidate => !ReferenceEquals(candidate, field)), cancellationToken).Any(candidate => candidate.Format == BibliographyFormat.EndNoteXml && string.Equals(candidate.Name, "year", StringComparison.OrdinalIgnoreCase));
    }

    private static void WriteUrls(XmlWriter writer, BibliographyItem item, BibliographyConversionReport report, string xmlNamespace, CancellationToken cancellationToken) {
        BibliographyNativeField[] additionalUrls = Cancellable(item.NativeFields, cancellationToken).Where(field => field.Format == BibliographyFormat.EndNoteXml && IsAdditionalUrlField(field, xmlNamespace)).ToArray();
        if (item.Url == null && additionalUrls.Length == 0) return;
        writer.WriteStartElement(null, "urls", xmlNamespace); writer.WriteStartElement(null, "related-urls", xmlNamespace);
        writer.WriteElementString(null, "url", xmlNamespace, item.Url == null ? string.Empty : SanitizeXml(item.Url));
        foreach (BibliographyNativeField field in Cancellable(additionalUrls, cancellationToken)) {
            if (TryWriteNativeField(writer, field, xmlNamespace)) report.Add("BIBCONV014", BibliographyDiagnosticSeverity.Information, "Preserved an additional EndNote XML related URL.", BibliographyConversionAction.PreservedExtension, item, "url");
            else report.Add("BIBCONV123", BibliographyDiagnosticSeverity.Warning, "An additional EndNote XML related URL is malformed and was omitted.", BibliographyConversionAction.Omitted, item, "url");
        }
        writer.WriteEndElement(); writer.WriteEndElement();
    }

    private static void WriteIdentifier(XmlWriter writer, BibliographyIdentifier identifier, string xmlNamespace) {
        if (string.Equals(identifier.Scheme, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase)) {
            writer.WriteStartElement(null, "isbn", xmlNamespace);
            if (!string.Equals(CodecMappings.InferSerialScheme(identifier.Value), identifier.Scheme, StringComparison.Ordinal)) writer.WriteAttributeString("type", identifier.Scheme);
            writer.WriteString(identifier.Value); writer.WriteEndElement();
        }
        else if (string.Equals(identifier.Scheme, "DOI", StringComparison.OrdinalIgnoreCase)) WriteElement(writer, "electronic-resource-num", identifier.Value, xmlNamespace);
        else if (string.Equals(identifier.Scheme, "accession", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "PMID", StringComparison.OrdinalIgnoreCase)) WriteElement(writer, "accession-num", identifier.Value, xmlNamespace);
    }

    private static bool TryWriteElement(XmlWriter writer, string xml) { try { XElement element = XElement.Parse(xml, LoadOptions.PreserveWhitespace); element.WriteTo(writer); return true; } catch (XmlException) { return false; } }
    private static bool TryWriteRootElement(XmlWriter writer, string xml, string rootNamespace) {
        try {
            XElement element = XElement.Parse(xml, LoadOptions.PreserveWhitespace);
            if (HasName(element, XNamespace.Get(rootNamespace), "records")) return false;
            element.WriteTo(writer);
            return true;
        } catch (XmlException) {
            return false;
        }
    }
    private static bool TryWriteRecordsElement(XmlWriter writer, string xml, string recordsNamespace) {
        try {
            XElement element = XElement.Parse(xml, LoadOptions.PreserveWhitespace);
            if (HasName(element, XNamespace.Get(recordsNamespace), "record")) return false;
            element.WriteTo(writer);
            return true;
        } catch (XmlException) {
            return false;
        }
    }
    private static void CaptureAttributes(XElement element, IList<BibliographyNativeEntry> nativeEntries, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        if (!HasElementMetadata(element)) return;
        string serialized = SerializeAttributes(element);
        foreach (XAttribute attribute in element.Attributes()) limits.AddValue(items, attribute.Value, GetOffset(element));
        limits.AddValue(items, serialized, GetOffset(element));
        nativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, AttributesEntryKind, serialized, element.Name.LocalName));
    }
    private static string SerializeAttributes(XElement element) =>
        new XElement(element.Name, element.Attributes().Select(static attribute => new XAttribute(attribute))).ToString(SaveOptions.DisableFormatting);
    private static bool HasElementMetadata(XElement element) => element.HasAttributes || element.Name.Namespace != XNamespace.None;
    private static string GetDocumentElementNamespace(BibliographyDocument document, string elementName, string fallback, CancellationToken cancellationToken) {
        BibliographyNativeEntry? entry = Cancellable(document.NativeEntries, cancellationToken).FirstOrDefault(candidate => IsAttributesEntry(candidate, elementName));
        return entry == null ? fallback : GetCarrierNamespace(entry.Value, fallback);
    }
    private static string GetRecordNamespace(BibliographyItem item, string fallback, CancellationToken cancellationToken) {
        BibliographyNativeField? field = Cancellable(item.NativeFields, cancellationToken).FirstOrDefault(candidate => candidate.Format == BibliographyFormat.EndNoteXml && string.Equals(candidate.Name, RecordAttributesFieldName, StringComparison.Ordinal));
        return field == null ? fallback : GetCarrierNamespace(field.Value, fallback);
    }
    private static string GetCarrierNamespace(string serializedCarrier, string fallback) {
        try { return XElement.Parse(serializedCarrier, LoadOptions.PreserveWhitespace).Name.NamespaceName; } catch (XmlException) { return fallback; }
    }
    private static void WriteDocumentAttributes(XmlWriter writer, BibliographyDocument document, string elementName, BibliographyConversionReport report, CancellationToken cancellationToken) {
        BibliographyNativeEntry? entry = Cancellable(document.NativeEntries, cancellationToken).FirstOrDefault(candidate => IsAttributesEntry(candidate, elementName));
        if (entry == null) return;
        if (TryWriteAttributes(writer, entry.Value)) report.Add("BIBCONV018", BibliographyDiagnosticSeverity.Information, $"Preserved EndNote XML attributes on '{elementName}'.", BibliographyConversionAction.PreservedExtension, field: elementName);
        else report.Add("BIBCONV131", BibliographyDiagnosticSeverity.Warning, $"EndNote XML attributes on '{elementName}' are malformed or conflicting and were omitted.", BibliographyConversionAction.Omitted, field: elementName);
    }
    internal static bool CoalescesRecordsContainerMetadata(BibliographyDocument document, CancellationToken cancellationToken) {
        int count = 0;
        foreach (BibliographyNativeEntry entry in document.NativeEntries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsAttributesEntry(entry, "records") && ++count > 1) return true;
        }
        return false;
    }
    private static void WriteRecordAttributes(XmlWriter writer, BibliographyItem item, BibliographyConversionReport report, CancellationToken cancellationToken) {
        foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken).Where(field => field.Format == BibliographyFormat.EndNoteXml && string.Equals(field.Name, RecordAttributesFieldName, StringComparison.Ordinal))) {
            if (TryWriteAttributes(writer, field.Value)) report.Add("BIBCONV019", BibliographyDiagnosticSeverity.Information, "Preserved EndNote XML record attributes.", BibliographyConversionAction.PreservedExtension, item, field.Name);
            else report.Add("BIBCONV132", BibliographyDiagnosticSeverity.Warning, "EndNote XML record attributes are malformed or conflicting and were omitted.", BibliographyConversionAction.Omitted, item, field.Name);
        }
    }
    private static bool TryWriteAttributes(XmlWriter writer, string serializedCarrier) {
        try {
            XElement carrier = XElement.Parse(serializedCarrier, LoadOptions.PreserveWhitespace);
            XAttribute[] attributes = carrier.Attributes().ToArray();
            var names = new HashSet<XName>();
            var namespacePrefixes = new HashSet<string>(StringComparer.Ordinal);
            foreach (XAttribute attribute in attributes) {
                if (attribute.IsNamespaceDeclaration) {
                    string declaredPrefix = attribute.Name.LocalName == "xmlns" ? string.Empty : attribute.Name.LocalName;
                    if (!namespacePrefixes.Add(declaredPrefix)) return false;
                } else if (!names.Add(attribute.Name)) return false;
            }
            foreach (XAttribute attribute in attributes.Where(static attribute => attribute.IsNamespaceDeclaration)) {
                if (attribute.Name.LocalName != "xmlns") writer.WriteAttributeString("xmlns", attribute.Name.LocalName, "http://www.w3.org/2000/xmlns/", attribute.Value);
            }
            foreach (XAttribute attribute in attributes.Where(static attribute => !attribute.IsNamespaceDeclaration)) {
                string? prefix = attribute.Name.Namespace == XNamespace.None ? null : carrier.GetPrefixOfNamespace(attribute.Name.Namespace);
                writer.WriteAttributeString(prefix, attribute.Name.LocalName, attribute.Name.NamespaceName, attribute.Value);
            }
            return true;
        } catch (Exception exception) when (exception is XmlException || exception is InvalidOperationException || exception is ArgumentException) {
            return false;
        }
    }
    private static bool IsAttributesEntry(BibliographyNativeEntry entry, string elementName) =>
        entry.Format == BibliographyFormat.EndNoteXml && string.Equals(entry.Kind, AttributesEntryKind, StringComparison.Ordinal) && string.Equals(entry.Name, elementName, StringComparison.OrdinalIgnoreCase);
    private static bool IsConsumedEndNoteEntry(BibliographyNativeEntry entry) =>
        entry.Format == BibliographyFormat.EndNoteXml && (entry.Kind == "element" || entry.Kind == RecordsElementEntryKind || IsAttributesEntry(entry, "xml") || IsAttributesEntry(entry, "records"));
    private static bool TryWriteNativeField(XmlWriter writer, BibliographyNativeField field, string xmlNamespace) {
        string? raw = field.UnmodifiedRawValue;
        if (raw != null) return TryWriteElement(writer, raw);
        if (HasInvalidXmlCharacters(field.Value)) return false;
        if (field.RawValue != null) return TryWriteEditedNativeElement(writer, field);
        try {
            XmlConvert.VerifyNCName(field.Name);
            writer.WriteElementString(null, field.Name, xmlNamespace, SanitizeXml(field.Value));
            return true;
        } catch (XmlException) {
            return false;
        }
    }
    private static bool TryWriteEditedNativeElement(XmlWriter writer, BibliographyNativeField field) {
        try {
            XElement original = XElement.Parse(field.RawValue!, LoadOptions.PreserveWhitespace);
            var edited = new XElement(original.Name, original.Attributes(), SanitizeXml(field.Value));
            edited.WriteTo(writer);
            return true;
        } catch (Exception exception) when (exception is XmlException || exception is InvalidOperationException || exception is ArgumentException) {
            return false;
        }
    }
    internal static bool EditedNativeFieldFlattensStructure(BibliographyNativeField field) {
        if (field.Format != BibliographyFormat.EndNoteXml || field.RawValue == null || field.UnmodifiedRawValue != null) return false;
        try {
            XElement original = XElement.Parse(field.RawValue, LoadOptions.PreserveWhitespace);
            return original.Nodes().Any(static node => !(node is XText));
        } catch (Exception exception) when (exception is XmlException || exception is InvalidOperationException || exception is ArgumentException) {
            return false;
        }
    }
    private static bool ConflictsWithTypedRecordElement(BibliographyItem item, BibliographyNativeField field, string xmlNamespace, CancellationToken cancellationToken) {
        if (string.Equals(field.Name, "periodical", StringComparison.OrdinalIgnoreCase)) return false;
        if (!KnownRecordElements.Contains(field.Name)) return false;
        if (CanPreserveUnownedKnownRecordContainer(item, field, xmlNamespace, cancellationToken)) return false;
        string? raw = field.UnmodifiedRawValue ?? field.RawValue;
        if (raw == null) return true;
        try {
            XElement element = XElement.Parse(raw, LoadOptions.PreserveWhitespace);
            if (IsRepeatableRecordElement(field.Name) && string.IsNullOrWhiteSpace(element.Value)) return false;
            return string.Equals(element.Name.NamespaceName, xmlNamespace, StringComparison.Ordinal);
        } catch (XmlException) {
            return true;
        }
    }
    private static bool CanPreserveUnownedKnownRecordContainer(BibliographyItem item, BibliographyNativeField field, string xmlNamespace, CancellationToken cancellationToken) {
        if (field.UnmodifiedRawValue == null) return false;
        try {
            XElement element = XElement.Parse(field.UnmodifiedRawValue, LoadOptions.PreserveWhitespace);
            if (!string.Equals(element.Name.NamespaceName, xmlNamespace, StringComparison.Ordinal) || !KnownRecordElements.Contains(element.Name.LocalName)) return false;
            switch (element.Name.LocalName.ToLowerInvariant()) {
                case "contributors": return item.Contributors.Count == 0;
                case "titles": return item.Title == null && item.ContainerTitle == null && item.CollectionTitle == null;
                case "periodical": return item.ContainerTitle == null;
                case "dates": return item.Dates.Count == 0;
                case "urls": return item.Url == null && !Cancellable(item.NativeFields, cancellationToken).Any(candidate => !ReferenceEquals(candidate, field) && IsAdditionalUrlField(candidate, xmlNamespace));
                case "keywords": return item.Keywords.Count == 0;
                default: return false;
            }
        } catch (XmlException) {
            return false;
        }
    }
    private static bool IsEmptyKnownRecordContainer(XElement element) {
        if (element.HasAttributes || element.Nodes().Any()) return false;
        switch (element.Name.LocalName.ToLowerInvariant()) {
            case "contributors": case "titles": case "periodical": case "dates": case "urls": case "keywords": return true;
            default: return false;
        }
    }
    private static bool IsAdditionalUrlField(BibliographyNativeField field, string xmlNamespace) {
        if (!string.Equals(field.Name, "url", StringComparison.OrdinalIgnoreCase)) return false;
        string? raw = field.UnmodifiedRawValue ?? field.RawValue;
        if (raw == null) return true;
        try {
            XElement element = XElement.Parse(raw, LoadOptions.PreserveWhitespace);
            return string.Equals(element.Name.LocalName, "url", StringComparison.OrdinalIgnoreCase) && string.Equals(element.Name.NamespaceName, xmlNamespace, StringComparison.Ordinal);
        } catch (XmlException) {
            return true;
        }
    }
    private static bool HasInvalidXmlCharacters(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) { index++; continue; }
            if (!XmlConvert.IsXmlChar(value[index])) return true;
        }
        return false;
    }

    private static int ValidateAggregateValueLengths(XElement element, IList<BibliographyItem> partial, BibliographyLimitGuard limits, bool checkCurrent) {
        int length = 0;
        foreach (XAttribute attribute in element.Attributes()) limits.AddValue(partial, attribute.Value, GetOffset(element));
        foreach (XNode node in element.Nodes()) {
            int nodeLength = node is XElement child ? ValidateAggregateValueLengths(child, partial, limits, ShouldCheckAggregateValue(child)) : node is XText text ? text.Value.Length : 0;
            if (nodeLength > int.MaxValue - length) throw new BibliographyLimitException("Maximum bibliography value length was exceeded.", partial, GetOffset(element));
            length += nodeLength;
        }
        if (checkCurrent) limits.CheckValueLength(partial, length, GetOffset(element));
        return length;
    }
    private static string SerializeBoundedElement(XElement element, IList<BibliographyItem> partial, BibliographyLimitGuard limits) {
        string serialized = element.ToString(SaveOptions.DisableFormatting);
        limits.CheckValueLength(partial, serialized, GetOffset(element));
        return serialized;
    }
    private static bool ShouldCheckAggregateValue(XElement element) {
        if (string.Equals(element.Parent?.Name.LocalName, "record", StringComparison.OrdinalIgnoreCase) && HasUnsupportedNestedContent(element)) return true;
        return !IsKnownContainer(element.Name.LocalName);
    }
    private static void WriteElement(XmlWriter writer, string name, string? value, string xmlNamespace) { if (value != null) writer.WriteElementString(null, name, xmlNamespace, SanitizeXml(value)); }
    private static string SanitizeXml(string value) {
        var builder = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) { builder.Append(value[index]).Append(value[++index]); continue; }
            builder.Append(XmlConvert.IsXmlChar(value[index]) ? value[index] : '\uFFFD');
        }
        return builder.ToString();
    }
    private static bool HasUnsupportedNestedContent(XElement element) {
        if (element.Attributes().Any(attribute => !IsKnownAttribute(element, attribute))) return true;
        foreach (XElement descendant in element.Descendants()) {
            if (descendant.Attributes().Any() || !IsKnownNestedElement(element, descendant)) return true;
        }
        foreach (XElement container in element.DescendantsAndSelf().Where(static candidate => IsKnownContainer(candidate.Name.LocalName))) {
            if (container.Nodes().Any(static node => node is XText text && !string.IsNullOrWhiteSpace(text.Value) || !(node is XElement) && !(node is XText))) return true;
        }
        return false;
    }
    private static bool IsKnownContainer(string name) {
        switch (name.ToLowerInvariant()) {
            case "xml": case "records": case "record": case "contributors": case "authors": case "secondary-authors": case "tertiary-authors": case "subsidiary-authors":
            case "titles": case "periodical": case "dates": case "pub-dates": case "urls": case "related-urls": case "keywords": return true;
            default: return false;
        }
    }
    private static bool IsKnownAttribute(XElement element, XAttribute attribute) {
        if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace != XNamespace.None) return false;
        if (string.Equals(element.Name.LocalName, "ref-type", StringComparison.OrdinalIgnoreCase) && string.Equals(attribute.Name.LocalName, "name", StringComparison.OrdinalIgnoreCase)) return true;
        return string.Equals(element.Name.LocalName, "isbn", StringComparison.OrdinalIgnoreCase) && string.Equals(attribute.Name.LocalName, "type", StringComparison.OrdinalIgnoreCase) &&
            (string.Equals(attribute.Value, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(attribute.Value, "ISSN", StringComparison.OrdinalIgnoreCase));
    }
    private static bool HasDuplicateKnownNestedContent(XElement element) {
        foreach (XElement parent in element.DescendantsAndSelf()) {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement child in parent.Elements()) {
                if (IsRepeatableNestedElement(child.Name.LocalName)) continue;
                if (!names.Add(child.Name.NamespaceName + "\0" + child.Name.LocalName)) return true;
            }
        }
        return false;
    }
    private static bool IsRepeatableRecordElement(string name) => string.Equals(name, "isbn", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "electronic-resource-num", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "accession-num", StringComparison.OrdinalIgnoreCase);
    private static bool IsRepeatableNestedElement(string name) => string.Equals(name, "author", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "url", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "keyword", StringComparison.OrdinalIgnoreCase);
    private static bool IsKnownNestedElement(XElement container, XElement descendant) {
        if (!HasNameInNamespace(descendant, container.Name.Namespace)) return false;
        string name = descendant.Name.LocalName.ToLowerInvariant();
        XElement? parent = descendant.Parent;
        switch (container.Name.LocalName.ToLowerInvariant()) {
            case "titles": return ReferenceEquals(parent, container) && (name == "title" || name == "secondary-title" || name == "tertiary-title");
            case "periodical": return ReferenceEquals(parent, container) && name == "full-title";
            case "contributors":
                if (ReferenceEquals(parent, container)) return IsContributorRoleElement(name);
                return name == "author" && parent != null && ReferenceEquals(parent.Parent, container) && IsContributorRoleElement(parent.Name.LocalName);
            case "dates":
                if (ReferenceEquals(parent, container)) return name == "year" || name == "pub-dates";
                return name == "date" && parent != null && ReferenceEquals(parent.Parent, container) && string.Equals(parent.Name.LocalName, "pub-dates", StringComparison.OrdinalIgnoreCase);
            case "urls":
                if (ReferenceEquals(parent, container)) return name == "related-urls";
                return name == "url" && parent != null && ReferenceEquals(parent.Parent, container) && string.Equals(parent.Name.LocalName, "related-urls", StringComparison.OrdinalIgnoreCase);
            case "keywords": return ReferenceEquals(parent, container) && name == "keyword";
            default: return false;
        }
    }
    private static bool IsContributorRoleElement(string name) {
        switch (name.ToLowerInvariant()) {
            case "authors": case "secondary-authors": case "tertiary-authors": case "subsidiary-authors": return true;
            default: return false;
        }
    }
    private static bool HasName(XElement element, XNamespace xmlNamespace, string localName) =>
        element.Name.Namespace == xmlNamespace && string.Equals(element.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase);
    private static bool HasNameInNamespace(XElement element, XNamespace xmlNamespace) => element.Name.Namespace == xmlNamespace;
    private static XElement? Child(XElement? parent, string name) {
        if (parent == null) return null;
        XNamespace xmlNamespace = parent.Name.Namespace;
        return parent.Elements().FirstOrDefault(element => HasName(element, xmlNamespace, name));
    }
    private static string Value(XElement? parent, string name) => Child(parent, name)?.Value ?? string.Empty;
    private static string? OptionalValue(XElement? parent, string name) => Child(parent, name)?.Value;
    private static string? FirstNonEmpty(params string?[] values) => values.FirstOrDefault(static value => !string.IsNullOrWhiteSpace(value)) ?? values.FirstOrDefault(static value => value != null);
    private static int GetOffset(XElement element) => element.Annotation<EndNoteSourceOffset>()?.Value ?? -1;
    private static void AddIdentifier(BibliographyItem item, string scheme, string value) { if (!string.IsNullOrWhiteSpace(value)) item.Identifiers.Add(new BibliographyIdentifier(scheme, value)); }
    private static IEnumerable<T> Cancellable<T>(IEnumerable<T> source, CancellationToken cancellationToken) {
        foreach (T value in source) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return value;
        }
    }
    private static void ParseAccessionIdentifier(BibliographyItem item, string value) {
        AddIdentifier(item, "accession", value);
    }
    private static BibliographyContributorRole RoleFromElement(string name) { switch (name.ToLowerInvariant()) { case "authors": return BibliographyContributorRole.Author; case "secondary-authors": return BibliographyContributorRole.Editor; case "tertiary-authors": return BibliographyContributorRole.CollectionEditor; case "subsidiary-authors": return BibliographyContributorRole.Translator; default: return BibliographyContributorRole.Other; } }
    private static string ElementFromRole(BibliographyContributorRole role) { switch (role) { case BibliographyContributorRole.Author: return "authors"; case BibliographyContributorRole.Editor: return "secondary-authors"; case BibliographyContributorRole.CollectionEditor: return "tertiary-authors"; case BibliographyContributorRole.Translator: return "subsidiary-authors"; default: return "subsidiary-authors"; } }
    private static string ToEndNoteType(BibliographyItemType type) { switch (type) { case BibliographyItemType.ArticleJournal: return "Journal Article"; case BibliographyItemType.Book: return "Book"; case BibliographyItemType.Chapter: return "Book Section"; case BibliographyItemType.PaperConference: return "Conference Paper"; case BibliographyItemType.Report: return "Report"; case BibliographyItemType.Thesis: return "Thesis"; case BibliographyItemType.WebPage: return "Web Page"; case BibliographyItemType.Patent: return "Patent"; default: return "Generic"; } }
    internal static bool CanPreserveNativeType(BibliographyFormat sourceFormat, BibliographyItem item) =>
        sourceFormat == BibliographyFormat.EndNoteXml && !string.IsNullOrWhiteSpace(item.NativeType) && CodecMappings.ParseType(item.NativeType) == item.Type;
    private static string OutputType(BibliographyFormat sourceFormat, BibliographyItem item) => CanPreserveNativeType(sourceFormat, item) ? item.NativeType! : ToEndNoteType(item.Type);
    private static int ToEndNoteNumber(BibliographyItemType type) { switch (type) { case BibliographyItemType.ArticleJournal: return 17; case BibliographyItemType.Book: return 6; case BibliographyItemType.Chapter: return 5; case BibliographyItemType.PaperConference: return 47; case BibliographyItemType.Report: return 27; case BibliographyItemType.Thesis: return 32; case BibliographyItemType.WebPage: return 12; case BibliographyItemType.Patent: return 21; default: return 13; } }

    private sealed class EncodingStringWriter : StringWriter {
        private readonly Encoding _encoding;
        internal EncodingStringWriter(StringBuilder builder, Encoding encoding) : base(builder, CultureInfo.InvariantCulture) => _encoding = encoding;
        public override Encoding Encoding => _encoding;
    }

}
