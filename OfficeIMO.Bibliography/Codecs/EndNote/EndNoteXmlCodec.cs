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

    internal static IList<BibliographyItem> Parse(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, IList<BibliographyNativeEntry> nativeEntries, out bool recordsRoot, out string? rootElementName, out string? recordsElementName, CancellationToken cancellationToken) {
        var items = new List<BibliographyItem>();
        recordsRoot = false;
        rootElementName = null;
        recordsElementName = null;
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
            if (root != null) foreach (XElement element in Cancellable(root.DescendantsAndSelf(), cancellationToken)) element.AddAnnotation(new EndNoteSourceOffset(offsets.GetOffset(element)));
            bool rootIsRecords = root != null && string.Equals(root.Name.LocalName, "records", StringComparison.OrdinalIgnoreCase);
            recordsRoot = rootIsRecords;
            rootElementName = root?.Name.LocalName;
            if (root != null) {
                CaptureAttributes(root, nativeEntries, items, limits, cancellationToken);
                if (!rootIsRecords) AddStructuralTextDiagnostic(root, diagnosticGuard, cancellationToken);
            }
            if (root != null && !rootIsRecords) foreach (XElement element in Cancellable(root.Elements(), cancellationToken)) {
                if (!HasName(element, root.Name.Namespace, "records")) {
                    ValidateAggregateValueLengths(element, items, limits, true, cancellationToken);
                    limits.AddValue(items, null, GetOffset(element));
                    nativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "element", SerializeBoundedElement(element, items, limits, cancellationToken), element.Name.LocalName));
                }
            }
            var containerList = new List<XElement>();
            if (root != null) {
                if (rootIsRecords) containerList.Add(root);
                else foreach (XElement element in Cancellable(root.Elements(), cancellationToken)) if (HasName(element, root.Name.Namespace, "records")) containerList.Add(element);
            }
            XElement[] containers = containerList.ToArray();
            recordsElementName = containers.FirstOrDefault()?.Name.LocalName;
            foreach (XElement container in Cancellable(containers, cancellationToken)) {
                AddStructuralTextDiagnostic(container, diagnosticGuard, cancellationToken);
                if (!ReferenceEquals(container, root)) CaptureAttributes(container, nativeEntries, items, limits, cancellationToken);
                foreach (XElement element in Cancellable(container.Elements(), cancellationToken)) {
                    if (!HasName(element, container.Name.Namespace, "record")) {
                        ValidateAggregateValueLengths(element, items, limits, true, cancellationToken);
                        limits.AddValue(items, null, GetOffset(element));
                        nativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, RecordsElementEntryKind, SerializeBoundedElement(element, items, limits, cancellationToken), element.Name.LocalName));
                    }
                }
            }
            foreach (XElement container in Cancellable(containers, cancellationToken)) {
                foreach (XElement record in Cancellable(container.Elements(), cancellationToken)) {
                    if (!HasName(record, container.Name.Namespace, "record")) continue;
                    BibliographyItem item = ParseRecord(record, items, limits, diagnosticGuard, cancellationToken);
                    items.Add(item);
                }
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
            string rootElementName = document.EndNoteRootElementName ?? (recordsRoot ? "records" : "xml");
            string outputNamespace = GetDocumentElementNamespace(document, rootElementName, string.Empty, cancellationToken);
            string? rootPrefix = GetDocumentElementPrefix(document, rootElementName, cancellationToken);
            writer.WriteStartDocument();
            if (!recordsRoot) {
                writer.WriteStartElement(rootPrefix, rootElementName, outputNamespace);
                WriteDocumentAttributes(writer, document, rootElementName, report, cancellationToken);
                foreach (BibliographyNativeEntry entry in Cancellable(document.NativeEntries, cancellationToken).Where(entry => entry.Format == BibliographyFormat.EndNoteXml && entry.Kind == "element")) {
                    if (TryWriteRootElement(writer, entry, outputNamespace)) report.Add("BIBCONV015", BibliographyDiagnosticSeverity.Information, $"Preserved document-level EndNote XML element '{entry.Name}'.", BibliographyConversionAction.PreservedExtension, field: entry.Name);
                    else report.Add("BIBCONV117", BibliographyDiagnosticSeverity.Warning, $"Document-level EndNote XML element '{entry.Name}' is malformed or reserved and was omitted.", BibliographyConversionAction.Omitted, field: entry.Name);
                }
            }
            string recordsElementName = document.EndNoteRecordsElementName ?? "records";
            string recordsNamespace = GetDocumentElementNamespace(document, recordsElementName, outputNamespace, cancellationToken);
            string? recordsPrefix = GetDocumentElementPrefix(document, recordsElementName, cancellationToken);
            writer.WriteStartElement(recordsPrefix, recordsElementName, recordsNamespace);
            WriteDocumentAttributes(writer, document, recordsElementName, report, cancellationToken);
            foreach (BibliographyNativeEntry entry in Cancellable(document.NativeEntries, cancellationToken).Where(entry => entry.Format == BibliographyFormat.EndNoteXml && entry.Kind == RecordsElementEntryKind)) {
                if (TryWriteRecordsElement(writer, entry, recordsNamespace)) report.Add("BIBCONV015", BibliographyDiagnosticSeverity.Information, $"Preserved EndNote XML records-container element '{entry.Name}'.", BibliographyConversionAction.PreservedExtension, field: entry.Name);
                else report.Add("BIBCONV117", BibliographyDiagnosticSeverity.Warning, $"EndNote XML records-container element '{entry.Name}' is malformed, reserved, or otherwise unsafe and was omitted.", BibliographyConversionAction.Omitted, field: entry.Name);
            }
            for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
                BibliographyItem item = document.Items[itemIndex];
                cancellationToken.ThrowIfCancellationRequested();
                string recordNamespace = GetRecordNamespace(item, recordsNamespace, cancellationToken);
                string? recordPrefix = GetRecordPrefix(item, cancellationToken);
                string recordElementName = GetRecordElementName(item, cancellationToken);
                writer.WriteStartElement(recordPrefix, recordElementName, recordNamespace);
                WriteRecordAttributes(writer, item, report, cancellationToken);
                WriteElement(writer, "rec-number", outputKeys[itemIndex], recordNamespace);
                writer.WriteStartElement(null, "ref-type", recordNamespace); writer.WriteAttributeString("name", SanitizeXml(OutputType(document.SourceFormat, item), cancellationToken)); writer.WriteString(ToEndNoteNumber(item.Type).ToString(CultureInfo.InvariantCulture)); writer.WriteEndElement();
                WriteContributors(writer, item, recordNamespace, cancellationToken); WriteTitles(writer, item, recordNamespace); WritePeriodical(writer, item, recordNamespace); WriteElement(writer, "pages", item.Pages, recordNamespace); WriteElement(writer, "volume", item.Volume, recordNamespace); WriteElement(writer, "number", item.Issue, recordNamespace);
                WriteElement(writer, "edition", item.Edition, recordNamespace); WriteElement(writer, "publisher", item.Publisher, recordNamespace); WriteElement(writer, "pub-location", item.PublisherPlace, recordNamespace);
                WriteElement(writer, "abstract", item.Abstract, recordNamespace); WriteElement(writer, "language", item.Language, recordNamespace); WriteDates(writer, item, report, recordNamespace, cancellationToken);
                foreach (BibliographyIdentifier identifier in Cancellable(item.Identifiers, cancellationToken)) WriteIdentifier(writer, identifier, recordNamespace, cancellationToken);
                bool hasAdditionalUrls = Cancellable(item.NativeFields, cancellationToken).Any(field => field.Format == BibliographyFormat.EndNoteXml && IsAdditionalUrlField(field, recordNamespace));
                bool wroteUrls = false;
                if (!hasAdditionalUrls) {
                    WriteUrls(writer, item, report, recordNamespace, cancellationToken);
                    wroteUrls = true;
                }
                if (item.Keywords.Count > 0) { writer.WriteStartElement(null, "keywords", recordNamespace); foreach (string keyword in Cancellable(item.Keywords, cancellationToken)) WriteElement(writer, "keyword", keyword, recordNamespace); writer.WriteEndElement(); }
                if (item.Notes.Count > 0) WriteElement(writer, "notes", string.Join("; ", Cancellable(item.Notes, cancellationToken)), recordNamespace);
                foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken)) {
                    if (field.Format == BibliographyFormat.EndNoteXml && string.Equals(field.Name, RecordAttributesFieldName, StringComparison.Ordinal)) {
                        continue;
                    } else if (field.Format == BibliographyFormat.EndNoteXml && IsAdditionalUrlField(field, recordNamespace)) {
                        if (!wroteUrls) {
                            WriteUrls(writer, item, report, recordNamespace, cancellationToken);
                            wroteUrls = true;
                        }
                        continue;
                    } else if (field.Format == BibliographyFormat.EndNoteXml && !ConflictsWithTypedRecordElement(item, field, recordNamespace, cancellationToken) && TryWriteNativeField(writer, field, recordNamespace, cancellationToken)) {
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
        foreach (BibliographyNativeEntry entry in Cancellable(document.NativeEntries, cancellationToken).Where(entry => !IsConsumedEndNoteEntry(document, entry))) report.Add("BIBCONV116", BibliographyDiagnosticSeverity.Warning, $"Document-level {entry.Format} entry '{entry.Kind}' cannot be represented in EndNote XML.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
        return builder.ToString();
    }

    internal static BibliographyItem ParseRecord(XElement record, IList<BibliographyItem> partial, BibliographyLimitGuard limits, BibliographyDiagnosticGuard diagnostics, CancellationToken cancellationToken) {
        ValidateAggregateValueLengths(record, partial, limits, false, cancellationToken);
        foreach (XElement leaf in Cancellable(record.Descendants(), cancellationToken)) if (!leaf.HasElements) limits.AddValue(partial, leaf.Value, GetOffset(leaf));
        XElement? refType = Child(record, "ref-type", cancellationToken);
        string type = refType?.Attribute("name")?.Value ?? refType?.Value ?? string.Empty;
        BibliographyItemType itemType = ParseEndNoteType(refType, type, out string nativeType);
        var item = new BibliographyItem { Key = Value(record, "rec-number", cancellationToken), NativeType = nativeType, Type = itemType };
        AddStructuralTextDiagnostic(record, diagnostics, cancellationToken, item.Key);
        if (refType?.Attribute("name") != null) {
            if (!int.TryParse(refType.Value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int nativeTypeNumber))
                diagnostics.Add(new BibliographyDiagnostic("BIBEND004", BibliographyDiagnosticSeverity.Warning, $"EndNote XML ref-type name '{type}' has nonnumeric code '{refType.Value}'.", GetOffset(refType), itemKey: item.Key, field: "ref-type"));
            else if (nativeTypeNumber != ToEndNoteNumber(item.Type))
                diagnostics.Add(new BibliographyDiagnostic("BIBEND004", BibliographyDiagnosticSeverity.Warning, $"EndNote XML ref-type name '{type}' conflicts with numeric code '{nativeTypeNumber}'.", GetOffset(refType), itemKey: item.Key, field: "ref-type"));
        }
        if (HasElementMetadata(record) || !string.Equals(record.Name.LocalName, "record", StringComparison.Ordinal)) {
            string recordMetadata = SerializeAttributes(record, cancellationToken);
            limits.AddValue(partial, recordMetadata, GetOffset(record));
            item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, RecordAttributesFieldName, recordMetadata));
        }
        XElement? titles = Child(record, "titles", cancellationToken); XElement? periodical = Child(record, "periodical", cancellationToken);
        string? secondaryTitle = OptionalValue(titles, "secondary-title", cancellationToken); string? periodicalTitle = OptionalValue(periodical, "full-title", cancellationToken);
        item.Title = OptionalValue(titles, "title", cancellationToken); item.ContainerTitle = secondaryTitle ?? periodicalTitle; item.CollectionTitle = OptionalValue(titles, "tertiary-title", cancellationToken);
        if (secondaryTitle != null) item.EndNoteFieldNames["container-title"] = "secondary-title";
        else if (periodicalTitle != null) item.EndNoteFieldNames["container-title"] = "periodical";
        bool retainedAdditionalPeriodical = periodical != null && secondaryTitle != null && periodicalTitle != null;
        item.Pages = OptionalValue(record, "pages", cancellationToken); item.Volume = OptionalValue(record, "volume", cancellationToken); item.Issue = OptionalValue(record, "number", cancellationToken); item.Edition = OptionalValue(record, "edition", cancellationToken);
        item.Publisher = OptionalValue(record, "publisher", cancellationToken); item.PublisherPlace = OptionalValue(record, "pub-location", cancellationToken); item.Abstract = OptionalValue(record, "abstract", cancellationToken); item.Language = OptionalValue(record, "language", cancellationToken);
        ParseContributors(item, Child(record, "contributors", cancellationToken), cancellationToken); ParseDates(item, Child(record, "dates", cancellationToken), cancellationToken);
        foreach (XElement identifier in Cancellable(record.Elements(), cancellationToken)) {
            if (!HasNameInNamespace(identifier, record.Name.Namespace)) continue;
            if (string.Equals(identifier.Name.LocalName, "isbn", StringComparison.OrdinalIgnoreCase)) {
                string? declaredScheme = identifier.Attribute("type")?.Value;
                string scheme = string.Equals(declaredScheme, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(declaredScheme, "ISSN", StringComparison.OrdinalIgnoreCase) ? declaredScheme! : CodecMappings.InferSerialScheme(identifier.Value);
                AddIdentifier(item, scheme, identifier.Value);
            } else if (string.Equals(identifier.Name.LocalName, "electronic-resource-num", StringComparison.OrdinalIgnoreCase)) AddIdentifier(item, "DOI", identifier.Value);
            else if (string.Equals(identifier.Name.LocalName, "accession-num", StringComparison.OrdinalIgnoreCase)) ParseAccessionIdentifier(item, identifier.Value);
        }
        XElement? urls = Child(record, "urls", cancellationToken); XElement? relatedUrlsContainer = Child(urls, "related-urls", cancellationToken);
        var relatedUrlList = new List<XElement>();
        if (relatedUrlsContainer != null) foreach (XElement element in Cancellable(relatedUrlsContainer.Elements(), cancellationToken)) if (HasName(element, record.Name.Namespace, "url")) relatedUrlList.Add(element);
        XElement[] relatedUrls = relatedUrlList.ToArray();
        string? primaryUrl = relatedUrls.FirstOrDefault()?.Value;
        item.Url = primaryUrl != null && (primaryUrl.Length > 0 || relatedUrls.Length == 1) ? primaryUrl : null;
        XElement? keywords = Child(record, "keywords", cancellationToken); if (keywords != null) foreach (XElement keyword in Cancellable(keywords.Elements(), cancellationToken)) if (HasName(keyword, keywords.Name.Namespace, "keyword")) item.Keywords.Add(keyword.Value);
        XElement? note = Child(record, "notes", cancellationToken); if (note != null) item.Notes.Add(note.Value);
        foreach (XElement element in Cancellable(record.Elements(), cancellationToken)) {
            bool knownRecordElement = HasNameInNamespace(element, record.Name.Namespace) && KnownRecordElements.Contains(element.Name.LocalName);
            bool repeatedSingleValue = knownRecordElement && !IsRepeatableRecordElement(element.Name.LocalName) && Cancellable(element.ElementsBeforeSelf(), cancellationToken).Any(previous => HasName(previous, record.Name.Namespace, element.Name.LocalName));
            if (!knownRecordElement) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, element.Name.LocalName, element.Value, SerializeBoundedElement(element, partial, limits, cancellationToken)));
            else if (ReferenceEquals(element, periodical) && retainedAdditionalPeriodical) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "periodical", element.Value, SerializeBoundedElement(element, partial, limits, cancellationToken)));
            else if (IsRepeatableRecordElement(element.Name.LocalName) && string.IsNullOrWhiteSpace(element.Value)) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, element.Name.LocalName, element.Value, SerializeBoundedElement(element, partial, limits, cancellationToken)));
            else if (IsEmptyKnownRecordContainer(element, cancellationToken) || HasEmptyKnownNestedContainer(element, cancellationToken) || repeatedSingleValue || HasUnsupportedNestedContent(element, cancellationToken) || HasDuplicateKnownNestedContent(element, cancellationToken)) item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.EndNoteXml, element.Name.LocalName, element.Value, SerializeBoundedElement(element, partial, limits, cancellationToken)));
            if (ReferenceEquals(element, urls)) foreach (XElement relatedUrl in Cancellable(relatedUrls.Skip(1), cancellationToken)) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "url", relatedUrl.Value, SerializeBoundedElement(relatedUrl, partial, limits, cancellationToken)));
        }
        if (string.IsNullOrWhiteSpace(item.Key)) diagnostics.Add(new BibliographyDiagnostic("BIBEND003", BibliographyDiagnosticSeverity.Warning, "EndNote XML record has no rec-number."));
        return item;
    }

    private static void ParseContributors(BibliographyItem item, XElement? contributors, CancellationToken cancellationToken) {
        if (contributors == null) return;
        foreach (XElement group in Cancellable(contributors.Elements(), cancellationToken)) {
            if (!HasNameInNamespace(group, contributors.Name.Namespace) || !IsContributorRoleElement(group.Name.LocalName)) continue;
            BibliographyContributorRole role = RoleFromElement(group.Name.LocalName);
            foreach (XElement value in Cancellable(group.Elements(), cancellationToken)) if (HasName(value, group.Name.Namespace, "author")) item.Contributors.Add(new BibliographyContributor(role, CodecMappings.ParseCommaName(value.Value)));
        }
    }

    private static void ParseDates(BibliographyItem item, XElement? dates, CancellationToken cancellationToken) {
        if (dates == null) return;
        XElement? yearElement = Child(dates, "year", cancellationToken);
        XElement? dateElement = Child(Child(dates, "pub-dates", cancellationToken), "date", cancellationToken);
        if (yearElement == null && dateElement == null) return;
        string year = yearElement?.Value ?? string.Empty; string pubDate = dateElement?.Value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(year) && string.IsNullOrWhiteSpace(pubDate)) {
            XElement retained = dateElement ?? yearElement!;
            var emptyDate = new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = retained.Value };
            if (yearElement != null) emptyDate.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.EndNoteXml, "year", year, yearElement.ToString(SaveOptions.DisableFormatting)));
            item.Dates.Add(emptyDate);
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
        else if (string.IsNullOrWhiteSpace(pubDate)) {
            parsedYear.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.EndNoteXml, "date", pubDate, dateElement.ToString(SaveOptions.DisableFormatting)));
            item.Dates.Add(parsedYear);
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
            if (TryWriteNativeField(writer, nativeYear, xmlNamespace, cancellationToken)) report.Add("BIBCONV014", BibliographyDiagnosticSeverity.Information, "Preserved a distinct EndNote XML year component.", BibliographyConversionAction.PreservedExtension, item, "dates.year");
            else report.Add("BIBCONV123", BibliographyDiagnosticSeverity.Warning, "A distinct EndNote XML year component is malformed and was omitted.", BibliographyConversionAction.Omitted, item, "dates.year");
        } else if (date.Year.HasValue) WriteElement(writer, "year", date.Year.Value.ToString(CultureInfo.InvariantCulture), xmlNamespace);
        string formatted = CodecMappings.FormatDate(date);
        BibliographyNativeField? nativePublicationDate = Cancellable(date.NativeFields, cancellationToken).FirstOrDefault(field => CanPreserveNativePublicationDateField(date, field, cancellationToken));
        writer.WriteStartElement(null, "pub-dates", xmlNamespace);
        if (nativePublicationDate != null) {
            if (TryWriteNativeField(writer, nativePublicationDate, xmlNamespace, cancellationToken)) report.Add("BIBCONV014", BibliographyDiagnosticSeverity.Information, "Preserved a distinct empty EndNote XML publication-date component.", BibliographyConversionAction.PreservedExtension, item, "dates.date");
            else report.Add("BIBCONV123", BibliographyDiagnosticSeverity.Warning, "A distinct empty EndNote XML publication-date component is malformed and was omitted.", BibliographyConversionAction.Omitted, item, "dates.date");
        } else WriteElement(writer, "date", formatted, xmlNamespace);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    internal static bool CanPreserveNativeDateField(BibliographyDate date, BibliographyNativeField field, CancellationToken cancellationToken = default) {
        if (field.Format != BibliographyFormat.EndNoteXml || !string.Equals(field.Name, "year", StringComparison.OrdinalIgnoreCase)) return false;
        BibliographyDate parsed = CodecMappings.ParseDate(BibliographyDateRole.Issued, field.Value);
        return !parsed.Year.HasValue && !Cancellable(date.NativeFields.TakeWhile(candidate => !ReferenceEquals(candidate, field)), cancellationToken).Any(candidate => candidate.Format == BibliographyFormat.EndNoteXml && string.Equals(candidate.Name, "year", StringComparison.OrdinalIgnoreCase));
    }
    internal static bool CanPreserveNativePublicationDateField(BibliographyDate date, BibliographyNativeField field, CancellationToken cancellationToken = default) {
        if (field.Format != BibliographyFormat.EndNoteXml || !string.Equals(field.Name, "date", StringComparison.OrdinalIgnoreCase) || field.UnmodifiedRawValue == null || !string.IsNullOrWhiteSpace(field.Value)) return false;
        if (date.Month.HasValue || date.Day.HasValue || date.EndYear.HasValue || date.EndMonth.HasValue || date.EndDay.HasValue || date.Literal != null) return false;
        return !Cancellable(date.NativeFields.TakeWhile(candidate => !ReferenceEquals(candidate, field)), cancellationToken).Any(candidate => candidate.Format == BibliographyFormat.EndNoteXml && string.Equals(candidate.Name, "date", StringComparison.OrdinalIgnoreCase));
    }

    private static void WriteUrls(XmlWriter writer, BibliographyItem item, BibliographyConversionReport report, string xmlNamespace, CancellationToken cancellationToken) {
        BibliographyNativeField[] additionalUrls = Cancellable(item.NativeFields, cancellationToken).Where(field => field.Format == BibliographyFormat.EndNoteXml && IsAdditionalUrlField(field, xmlNamespace)).ToArray();
        if (item.Url == null && additionalUrls.Length == 0) return;
        writer.WriteStartElement(null, "urls", xmlNamespace); writer.WriteStartElement(null, "related-urls", xmlNamespace);
        writer.WriteElementString(null, "url", xmlNamespace, item.Url == null ? string.Empty : SanitizeXml(item.Url, cancellationToken));
        foreach (BibliographyNativeField field in Cancellable(additionalUrls, cancellationToken)) {
            if (TryWriteNativeField(writer, field, xmlNamespace, cancellationToken)) report.Add("BIBCONV014", BibliographyDiagnosticSeverity.Information, "Preserved an additional EndNote XML related URL.", BibliographyConversionAction.PreservedExtension, item, "url");
            else report.Add("BIBCONV123", BibliographyDiagnosticSeverity.Warning, "An additional EndNote XML related URL is malformed and was omitted.", BibliographyConversionAction.Omitted, item, "url");
        }
        writer.WriteEndElement(); writer.WriteEndElement();
    }

    private static void WriteIdentifier(XmlWriter writer, BibliographyIdentifier identifier, string xmlNamespace, CancellationToken cancellationToken) {
        if (string.Equals(identifier.Scheme, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase)) {
            writer.WriteStartElement(null, "isbn", xmlNamespace);
            if (!string.Equals(CodecMappings.InferSerialScheme(identifier.Value), identifier.Scheme, StringComparison.Ordinal)) writer.WriteAttributeString("type", SanitizeXml(identifier.Scheme, cancellationToken));
            writer.WriteString(SanitizeXml(identifier.Value, cancellationToken)); writer.WriteEndElement();
        }
        else if (string.Equals(identifier.Scheme, "DOI", StringComparison.OrdinalIgnoreCase)) WriteElement(writer, "electronic-resource-num", identifier.Value, xmlNamespace, cancellationToken);
        else if (string.Equals(identifier.Scheme, "accession", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "PMID", StringComparison.OrdinalIgnoreCase)) WriteElement(writer, "accession-num", identifier.Value, xmlNamespace, cancellationToken);
    }

    private static bool TryWriteElement(XmlWriter writer, string xml, CancellationToken cancellationToken = default) {
        try {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = XElement.Parse(xml, LoadOptions.PreserveWhitespace);
            cancellationToken.ThrowIfCancellationRequested();
            element.WriteTo(writer);
            cancellationToken.ThrowIfCancellationRequested();
            return true;
        } catch (XmlException) {
            return false;
        }
    }
    private static bool TryWriteRootElement(XmlWriter writer, BibliographyNativeEntry entry, string rootNamespace) {
        try {
            XElement element = XElement.Parse(entry.Value, LoadOptions.PreserveWhitespace);
            if (!string.Equals(entry.Name, element.Name.LocalName, StringComparison.Ordinal)) return false;
            if (HasName(element, XNamespace.Get(rootNamespace), "records")) return false;
            element.WriteTo(writer);
            return true;
        } catch (XmlException) {
            return false;
        }
    }
    private static bool TryWriteRecordsElement(XmlWriter writer, BibliographyNativeEntry entry, string recordsNamespace) {
        try {
            XElement element = XElement.Parse(entry.Value, LoadOptions.PreserveWhitespace);
            if (!string.Equals(entry.Name, element.Name.LocalName, StringComparison.Ordinal)) return false;
            if (HasName(element, XNamespace.Get(recordsNamespace), "record")) return false;
            element.WriteTo(writer);
            return true;
        } catch (XmlException) {
            return false;
        }
    }
    private static void CaptureAttributes(XElement element, IList<BibliographyNativeEntry> nativeEntries, IList<BibliographyItem> items, BibliographyLimitGuard limits, CancellationToken cancellationToken) {
        if (!HasElementMetadata(element)) return;
        string serialized = SerializeAttributes(element, cancellationToken);
        foreach (XAttribute attribute in Cancellable(element.Attributes(), cancellationToken)) limits.AddValue(items, attribute.Value, GetOffset(element));
        limits.AddValue(items, serialized, GetOffset(element));
        nativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, AttributesEntryKind, serialized, element.Name.LocalName));
    }
    private static string SerializeAttributes(XElement element, CancellationToken cancellationToken) {
        var attributes = Cancellable(element.Attributes(), cancellationToken).Select(static attribute => new XAttribute(attribute)).ToList();
        AddInheritedNamespaceDeclaration(element, element.Name.Namespace, attributes);
        foreach (XAttribute attribute in Cancellable(element.Attributes(), cancellationToken)) {
            if (!attribute.IsNamespaceDeclaration) AddInheritedNamespaceDeclaration(element, attribute.Name.Namespace, attributes);
        }
        return new XElement(element.Name, attributes).ToString(SaveOptions.DisableFormatting);
    }
    private static void AddInheritedNamespaceDeclaration(XElement element, XNamespace xmlNamespace, IList<XAttribute> attributes) {
        if (xmlNamespace == XNamespace.None || xmlNamespace == XNamespace.Xml) return;
        string? prefix = element.GetPrefixOfNamespace(xmlNamespace);
        if (prefix == null || attributes.Any(attribute => IsNamespaceDeclaration(attribute, prefix))) return;
        attributes.Add(prefix.Length == 0
            ? new XAttribute("xmlns", xmlNamespace.NamespaceName)
            : new XAttribute(XNamespace.Xmlns + prefix, xmlNamespace.NamespaceName));
    }
    private static bool IsNamespaceDeclaration(XAttribute attribute, string prefix) =>
        attribute.IsNamespaceDeclaration && string.Equals(attribute.Name.LocalName == "xmlns" ? string.Empty : attribute.Name.LocalName, prefix, StringComparison.Ordinal);
    private static bool HasElementMetadata(XElement element) => element.HasAttributes || element.Name.Namespace != XNamespace.None;
    private static string GetDocumentElementNamespace(BibliographyDocument document, string elementName, string fallback, CancellationToken cancellationToken) {
        BibliographyNativeEntry? entry = Cancellable(document.NativeEntries, cancellationToken).FirstOrDefault(candidate => IsAttributesEntry(candidate, elementName));
        return entry == null ? fallback : GetCarrierNamespace(entry.Value, fallback);
    }
    private static string? GetDocumentElementPrefix(BibliographyDocument document, string elementName, CancellationToken cancellationToken) {
        BibliographyNativeEntry? entry = Cancellable(document.NativeEntries, cancellationToken).FirstOrDefault(candidate => IsAttributesEntry(candidate, elementName));
        return entry == null ? null : GetCarrierPrefix(entry.Value);
    }
    private static string GetRecordNamespace(BibliographyItem item, string fallback, CancellationToken cancellationToken) {
        BibliographyNativeField? field = Cancellable(item.NativeFields, cancellationToken).FirstOrDefault(candidate => candidate.Format == BibliographyFormat.EndNoteXml && string.Equals(candidate.Name, RecordAttributesFieldName, StringComparison.Ordinal));
        return field == null ? fallback : GetCarrierNamespace(field.Value, fallback);
    }
    private static string? GetRecordPrefix(BibliographyItem item, CancellationToken cancellationToken) {
        BibliographyNativeField? field = Cancellable(item.NativeFields, cancellationToken).FirstOrDefault(candidate => candidate.Format == BibliographyFormat.EndNoteXml && string.Equals(candidate.Name, RecordAttributesFieldName, StringComparison.Ordinal));
        return field == null ? null : GetCarrierPrefix(field.Value);
    }
    private static string GetRecordElementName(BibliographyItem item, CancellationToken cancellationToken) {
        BibliographyNativeField? field = Cancellable(item.NativeFields, cancellationToken).FirstOrDefault(candidate => candidate.Format == BibliographyFormat.EndNoteXml && string.Equals(candidate.Name, RecordAttributesFieldName, StringComparison.Ordinal));
        if (field == null) return "record";
        try {
            string localName = XElement.Parse(field.Value, LoadOptions.PreserveWhitespace).Name.LocalName;
            return string.Equals(localName, "record", StringComparison.OrdinalIgnoreCase) ? localName : "record";
        } catch (XmlException) {
            return "record";
        }
    }
    private static string GetCarrierNamespace(string serializedCarrier, string fallback) {
        try { return XElement.Parse(serializedCarrier, LoadOptions.PreserveWhitespace).Name.NamespaceName; } catch (XmlException) { return fallback; }
    }
    private static string? GetCarrierPrefix(string serializedCarrier) {
        try {
            XElement carrier = XElement.Parse(serializedCarrier, LoadOptions.PreserveWhitespace);
            string? prefix = carrier.GetPrefixOfNamespace(carrier.Name.Namespace);
            return string.IsNullOrEmpty(prefix) ? null : prefix;
        } catch (XmlException) {
            return null;
        }
    }
    private static void WriteDocumentAttributes(XmlWriter writer, BibliographyDocument document, string elementName, BibliographyConversionReport report, CancellationToken cancellationToken) {
        BibliographyNativeEntry? entry = Cancellable(document.NativeEntries, cancellationToken).FirstOrDefault(candidate => IsAttributesEntry(candidate, elementName));
        if (entry == null) return;
        if (TryWriteAttributes(writer, entry.Value)) report.Add("BIBCONV018", BibliographyDiagnosticSeverity.Information, $"Preserved EndNote XML attributes on '{elementName}'.", BibliographyConversionAction.PreservedExtension, field: elementName);
        else report.Add("BIBCONV131", BibliographyDiagnosticSeverity.Warning, $"EndNote XML attributes on '{elementName}' are malformed or conflicting and were omitted.", BibliographyConversionAction.Omitted, field: elementName);
    }
    internal static bool CoalescesRecordsContainerMetadata(BibliographyDocument document, CancellationToken cancellationToken) {
        return HasDuplicateDocumentAttributes(document, "records", cancellationToken);
    }
    internal static bool HasDuplicateDocumentAttributes(BibliographyDocument document, string elementName, CancellationToken cancellationToken) {
        int count = 0;
        foreach (BibliographyNativeEntry entry in document.NativeEntries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsAttributesEntry(entry, elementName) && ++count > 1) return true;
        }
        return false;
    }
    private static void WriteRecordAttributes(XmlWriter writer, BibliographyItem item, BibliographyConversionReport report, CancellationToken cancellationToken) {
        bool wroteAttributes = false;
        foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken).Where(field => field.Format == BibliographyFormat.EndNoteXml && string.Equals(field.Name, RecordAttributesFieldName, StringComparison.Ordinal))) {
            if (wroteAttributes) {
                report.Add("BIBCONV249", BibliographyDiagnosticSeverity.Warning, "Additional EndNote XML record-attribute metadata was omitted to keep one source-preserving carrier.", BibliographyConversionAction.Omitted, item, field.Name);
                continue;
            }
            wroteAttributes = true;
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
            foreach (XAttribute attribute in attributes) {
                if (attribute.IsNamespaceDeclaration) {
                    string declaredPrefix = attribute.Name.LocalName == "xmlns" ? string.Empty : attribute.Name.LocalName;
                    string? elementPrefix = carrier.GetPrefixOfNamespace(carrier.Name.Namespace);
                    if (string.Equals(attribute.Value, carrier.Name.NamespaceName, StringComparison.Ordinal) &&
                        string.Equals(declaredPrefix, elementPrefix ?? string.Empty, StringComparison.Ordinal)) continue;
                    writer.WriteAttributeString("xmlns", declaredPrefix, "http://www.w3.org/2000/xmlns/", attribute.Value);
                } else {
                    string? prefix = attribute.Name.Namespace == XNamespace.None ? null : carrier.GetPrefixOfNamespace(attribute.Name.Namespace);
                    writer.WriteAttributeString(prefix, attribute.Name.LocalName, attribute.Name.NamespaceName, attribute.Value);
                }
            }
            return true;
        } catch (Exception exception) when (exception is XmlException || exception is InvalidOperationException || exception is ArgumentException) {
            return false;
        }
    }
    private static bool IsAttributesEntry(BibliographyNativeEntry entry, string elementName) =>
        entry.Format == BibliographyFormat.EndNoteXml && string.Equals(entry.Kind, AttributesEntryKind, StringComparison.Ordinal) && string.Equals(entry.Name, elementName, StringComparison.OrdinalIgnoreCase);
    private static bool IsConsumedEndNoteEntry(BibliographyDocument document, BibliographyNativeEntry entry) =>
        entry.Format == BibliographyFormat.EndNoteXml && (entry.Kind == "element" && !document.EndNoteRecordsRoot || entry.Kind == RecordsElementEntryKind ||
            IsAttributesEntry(entry, document.EndNoteRootElementName ?? (document.EndNoteRecordsRoot ? "records" : "xml")) ||
            IsAttributesEntry(entry, document.EndNoteRecordsElementName ?? "records"));
    private static bool TryWriteNativeField(XmlWriter writer, BibliographyNativeField field, string xmlNamespace, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string? raw = field.UnmodifiedRawValue;
        if (raw != null) return TryWriteElement(writer, raw, cancellationToken);
        if (HasInvalidXmlCharacters(field.Value, cancellationToken)) return false;
        if (field.RawValue != null) return TryWriteEditedNativeElement(writer, field, cancellationToken);
        try {
            XmlConvert.VerifyNCName(field.Name);
            writer.WriteElementString(null, field.Name, xmlNamespace, SanitizeXml(field.Value, cancellationToken));
            return true;
        } catch (XmlException) {
            return false;
        }
    }
    private static bool TryWriteEditedNativeElement(XmlWriter writer, BibliographyNativeField field, CancellationToken cancellationToken) {
        try {
            cancellationToken.ThrowIfCancellationRequested();
            XElement original = XElement.Parse(field.RawValue!, LoadOptions.PreserveWhitespace);
            cancellationToken.ThrowIfCancellationRequested();
            var edited = new XElement(original.Name, original.Attributes(), SanitizeXml(field.Value, cancellationToken));
            edited.WriteTo(writer);
            cancellationToken.ThrowIfCancellationRequested();
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
        if (string.Equals(field.Name, "periodical", StringComparison.OrdinalIgnoreCase) && item.ContainerTitle != null &&
            item.EndNoteFieldNames.TryGetValue("container-title", out string? source) && !string.Equals(source, "periodical", StringComparison.OrdinalIgnoreCase)) return false;
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
            if (ContainsBindableTypedContent(element, cancellationToken)) return false;
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
    private static bool ContainsBindableTypedContent(XElement element, CancellationToken cancellationToken) {
        XNamespace xmlNamespace = element.Name.Namespace;
        switch (element.Name.LocalName.ToLowerInvariant()) {
            case "contributors":
                return Cancellable(element.Elements(), cancellationToken).Any(group => HasNameInNamespace(group, xmlNamespace) && IsContributorRoleElement(group.Name.LocalName) &&
                    Cancellable(group.Elements(), cancellationToken).Any(author => HasName(author, xmlNamespace, "author")));
            case "titles":
                return Cancellable(element.Elements(), cancellationToken).Any(child => HasNameInNamespace(child, xmlNamespace) &&
                    (string.Equals(child.Name.LocalName, "title", StringComparison.OrdinalIgnoreCase) || string.Equals(child.Name.LocalName, "secondary-title", StringComparison.OrdinalIgnoreCase) || string.Equals(child.Name.LocalName, "tertiary-title", StringComparison.OrdinalIgnoreCase)));
            case "periodical":
                return Cancellable(element.Elements(), cancellationToken).Any(child => HasName(child, xmlNamespace, "full-title"));
            case "dates":
                return Cancellable(element.Elements(), cancellationToken).Any(child => HasName(child, xmlNamespace, "year") || HasName(child, xmlNamespace, "pub-dates") &&
                    Cancellable(child.Elements(), cancellationToken).Any(date => HasName(date, xmlNamespace, "date")));
            case "urls":
                return Cancellable(element.Elements(), cancellationToken).Any(child => HasName(child, xmlNamespace, "related-urls") &&
                    Cancellable(child.Elements(), cancellationToken).Any(url => HasName(url, xmlNamespace, "url")));
            case "keywords":
                return Cancellable(element.Elements(), cancellationToken).Any(child => HasName(child, xmlNamespace, "keyword"));
            default:
                return false;
        }
    }
    private static bool IsEmptyKnownRecordContainer(XElement element, CancellationToken cancellationToken = default) {
        if (element.HasAttributes || Cancellable(element.Nodes(), cancellationToken).Any(static node => !(node is XText text) || !string.IsNullOrWhiteSpace(text.Value))) return false;
        switch (element.Name.LocalName.ToLowerInvariant()) {
            case "contributors": case "titles": case "periodical": case "dates": case "urls": case "keywords": return true;
            default: return false;
        }
    }
    private static bool HasEmptyKnownNestedContainer(XElement element, CancellationToken cancellationToken = default) {
        foreach (XElement descendant in Cancellable(element.Descendants(), cancellationToken)) {
            if (!IsKnownContainer(descendant.Name.LocalName) || descendant.HasAttributes) continue;
            if (!Cancellable(descendant.Nodes(), cancellationToken).Any(static node => !(node is XText text) || !string.IsNullOrWhiteSpace(text.Value))) return true;
        }
        return false;
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
    internal static bool HasInvalidXmlCharacters(string value, CancellationToken cancellationToken = default) {
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) { index++; continue; }
            if (!XmlConvert.IsXmlChar(value[index])) return true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return false;
    }

    private static int ValidateAggregateValueLengths(XElement element, IList<BibliographyItem> partial, BibliographyLimitGuard limits, bool checkCurrent, CancellationToken cancellationToken = default) {
        int length = 0;
        foreach (XAttribute attribute in Cancellable(element.Attributes(), cancellationToken)) limits.AddValue(partial, attribute.Value, GetOffset(element));
        foreach (XNode node in Cancellable(element.Nodes(), cancellationToken)) {
            int nodeLength = node is XElement child ? ValidateAggregateValueLengths(child, partial, limits, ShouldCheckAggregateValue(child, cancellationToken), cancellationToken) : node is XText text ? text.Value.Length : 0;
            if (nodeLength > int.MaxValue - length) throw new BibliographyLimitException("Maximum bibliography value length was exceeded.", partial, GetOffset(element));
            length += nodeLength;
        }
        if (checkCurrent) limits.CheckValueLength(partial, length, GetOffset(element));
        return length;
    }
    private static string SerializeBoundedElement(XElement element, IList<BibliographyItem> partial, BibliographyLimitGuard limits, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string serialized = element.ToString(SaveOptions.DisableFormatting);
        cancellationToken.ThrowIfCancellationRequested();
        limits.CheckValueLength(partial, serialized, GetOffset(element));
        return serialized;
    }
    private static bool ShouldCheckAggregateValue(XElement element, CancellationToken cancellationToken = default) {
        if (string.Equals(element.Parent?.Name.LocalName, "record", StringComparison.OrdinalIgnoreCase) && HasUnsupportedNestedContent(element, cancellationToken)) return true;
        return !IsKnownContainer(element.Name.LocalName);
    }
    private static void WriteElement(XmlWriter writer, string name, string? value, string xmlNamespace, CancellationToken cancellationToken = default) { if (value != null) writer.WriteElementString(null, name, xmlNamespace, SanitizeXml(value, cancellationToken)); }
    internal static string SanitizeXml(string value, CancellationToken cancellationToken = default) {
        var builder = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (char.IsHighSurrogate(value[index])) {
                if (index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) builder.Append(value[index]).Append(value[++index]);
                else builder.Append('\uFFFD');
                continue;
            }
            if (char.IsLowSurrogate(value[index])) { builder.Append('\uFFFD'); continue; }
            builder.Append(XmlConvert.IsXmlChar(value[index]) ? value[index] : '\uFFFD');
        }
        cancellationToken.ThrowIfCancellationRequested();
        return builder.ToString();
    }
    private static bool HasUnsupportedNestedContent(XElement element, CancellationToken cancellationToken = default) {
        if (Cancellable(element.Attributes(), cancellationToken).Any(attribute => !IsKnownAttribute(element, attribute))) return true;
        foreach (XElement descendant in Cancellable(element.Descendants(), cancellationToken)) {
            if (Cancellable(descendant.Attributes(), cancellationToken).Any() || !IsKnownNestedElement(element, descendant)) return true;
        }
        foreach (XElement container in Cancellable(element.DescendantsAndSelf(), cancellationToken)) {
            if (!IsKnownContainer(container.Name.LocalName)) continue;
            if (Cancellable(container.Nodes(), cancellationToken).Any(static node => node is XText text && !string.IsNullOrWhiteSpace(text.Value) || !(node is XElement) && !(node is XText))) return true;
        }
        return false;
    }
    private static void AddStructuralTextDiagnostic(XElement element, BibliographyDiagnosticGuard diagnostics, CancellationToken cancellationToken, string? itemKey = null) {
        if (!Cancellable(element.Nodes(), cancellationToken).Any(static node => node is XText text && !string.IsNullOrWhiteSpace(text.Value))) return;
        diagnostics.Add(new BibliographyDiagnostic("BIBEND005", BibliographyDiagnosticSeverity.Warning,
            $"EndNote XML structural element '{element.Name.LocalName}' contains direct text that canonical output cannot retain.",
            GetOffset(element), itemKey: itemKey, field: element.Name.LocalName));
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
        if (string.Equals(element.Name.LocalName, "ref-type", StringComparison.OrdinalIgnoreCase) && string.Equals(attribute.Name.LocalName, "name", StringComparison.Ordinal)) return true;
        return string.Equals(element.Name.LocalName, "isbn", StringComparison.OrdinalIgnoreCase) && string.Equals(attribute.Name.LocalName, "type", StringComparison.Ordinal) &&
            (string.Equals(attribute.Value, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(attribute.Value, "ISSN", StringComparison.OrdinalIgnoreCase));
    }
    private static bool HasDuplicateKnownNestedContent(XElement element, CancellationToken cancellationToken = default) {
        foreach (XElement parent in Cancellable(element.DescendantsAndSelf(), cancellationToken)) {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement child in Cancellable(parent.Elements(), cancellationToken)) {
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
    private static XElement? Child(XElement? parent, string name, CancellationToken cancellationToken = default) {
        if (parent == null) return null;
        XNamespace xmlNamespace = parent.Name.Namespace;
        return Cancellable(parent.Elements(), cancellationToken).FirstOrDefault(element => HasName(element, xmlNamespace, name));
    }
    private static string Value(XElement? parent, string name, CancellationToken cancellationToken = default) => Child(parent, name, cancellationToken)?.Value ?? string.Empty;
    private static string? OptionalValue(XElement? parent, string name, CancellationToken cancellationToken = default) => Child(parent, name, cancellationToken)?.Value;
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
        sourceFormat == BibliographyFormat.EndNoteXml && !string.IsNullOrWhiteSpace(item.NativeType) && !int.TryParse(item.NativeType!.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out _) && CodecMappings.ParseType(item.NativeType) == item.Type;
    private static string OutputType(BibliographyFormat sourceFormat, BibliographyItem item) => CanPreserveNativeType(sourceFormat, item) ? item.NativeType! : ToEndNoteType(item.Type);
    private static int ToEndNoteNumber(BibliographyItemType type) { switch (type) { case BibliographyItemType.ArticleJournal: return 17; case BibliographyItemType.Book: return 6; case BibliographyItemType.Chapter: return 5; case BibliographyItemType.PaperConference: return 47; case BibliographyItemType.Report: return 27; case BibliographyItemType.Thesis: return 32; case BibliographyItemType.WebPage: return 12; case BibliographyItemType.Patent: return 21; default: return 13; } }
    private static BibliographyItemType ParseEndNoteType(XElement? refType, string sourceType, out string nativeType) {
        if (refType?.Attribute("name") != null) { nativeType = sourceType; return CodecMappings.ParseType(sourceType); }
        if (int.TryParse(sourceType.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int number) && TryParseEndNoteNumber(number, out BibliographyItemType parsed)) {
            nativeType = ToEndNoteType(parsed);
            return parsed;
        }
        nativeType = sourceType;
        return CodecMappings.ParseType(sourceType);
    }
    private static bool TryParseEndNoteNumber(int number, out BibliographyItemType type) {
        switch (number) {
            case 17: type = BibliographyItemType.ArticleJournal; return true;
            case 6: type = BibliographyItemType.Book; return true;
            case 5: type = BibliographyItemType.Chapter; return true;
            case 47: type = BibliographyItemType.PaperConference; return true;
            case 27: type = BibliographyItemType.Report; return true;
            case 32: type = BibliographyItemType.Thesis; return true;
            case 12: type = BibliographyItemType.WebPage; return true;
            case 21: type = BibliographyItemType.Patent; return true;
            case 13: type = BibliographyItemType.Document; return true;
            default: type = BibliographyItemType.Unknown; return false;
        }
    }

    private sealed class EncodingStringWriter : StringWriter {
        private readonly Encoding _encoding;
        internal EncodingStringWriter(StringBuilder builder, Encoding encoding) : base(builder, CultureInfo.InvariantCulture) => _encoding = encoding;
        public override Encoding Encoding => _encoding;
    }

}
