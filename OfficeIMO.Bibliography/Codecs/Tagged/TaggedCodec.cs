namespace OfficeIMO.Bibliography;

internal static class TaggedCodec {
    internal static IList<BibliographyItem> ParseRis(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, CancellationToken cancellationToken) =>
        Parse(source, BibliographyFormat.Ris, options, diagnostics, cancellationToken);

    internal static IList<BibliographyItem> ParseNbib(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, CancellationToken cancellationToken) =>
        Parse(source, BibliographyFormat.Nbib, options, diagnostics, cancellationToken);

    internal static string WriteRis(BibliographyDocument document, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        var builder = new StringBuilder();
        string[] outputKeys = CodecMappings.OutputKeys(document.Items, BibliographyFormat.Ris, cancellationToken);
        for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
            BibliographyItem item = document.Items[itemIndex];
            cancellationToken.ThrowIfCancellationRequested();
            WriteTag(builder, "TY", CanPreserveNativeType(document.SourceFormat, item) ? item.NativeType : CodecMappings.ToRisType(item.Type), options.LineEnding, cancellationToken);
            WriteTag(builder, "ID", outputKeys[itemIndex], options.LineEnding, cancellationToken);
            WriteTag(builder, TaggedOutputTag(item, BibliographyFormat.Ris, "title", "TI"), item.Title, options.LineEnding, cancellationToken);
            WriteTag(builder, TaggedOutputTag(item, BibliographyFormat.Ris, "container-title", "T2"), item.ContainerTitle, options.LineEnding, cancellationToken);
            foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken)) {
                if (contributor.Role == BibliographyContributorRole.Author) WriteTag(builder, ContributorTag(item, contributor, "AU", "A1"), CodecMappings.FormatName(contributor.Name), options.LineEnding, cancellationToken);
                else if (contributor.Role == BibliographyContributorRole.Editor) WriteTag(builder, ContributorTag(item, contributor, "ED", "A2"), CodecMappings.FormatName(contributor.Name), options.LineEnding, cancellationToken);
            }
            WriteDateTags(builder, item, options.LineEnding, "PY", "Y2", cancellationToken);
            WriteTag(builder, "PB", item.Publisher, options.LineEnding, cancellationToken); WriteTag(builder, "CY", item.PublisherPlace, options.LineEnding, cancellationToken); WriteTag(builder, "ET", item.Edition, options.LineEnding, cancellationToken);
            WriteTag(builder, "VL", item.Volume, options.LineEnding, cancellationToken); WriteTag(builder, "IS", item.Issue, options.LineEnding, cancellationToken); WriteRisPages(builder, item, options.LineEnding, cancellationToken);
            WriteTag(builder, TaggedOutputTag(item, BibliographyFormat.Ris, "abstract", "AB"), item.Abstract, options.LineEnding, cancellationToken); WriteTag(builder, "LA", item.Language, options.LineEnding, cancellationToken); WriteTag(builder, TaggedOutputTag(item, BibliographyFormat.Ris, "url", "UR"), item.Url, options.LineEnding, cancellationToken);
            foreach (BibliographyIdentifier identifier in Cancellable(item.Identifiers, cancellationToken)) WriteRisIdentifier(builder, identifier, options.LineEnding, cancellationToken);
            foreach (string keyword in Cancellable(item.Keywords, cancellationToken)) WriteTag(builder, "KW", keyword, options.LineEnding, cancellationToken);
            foreach (string note in Cancellable(item.Notes, cancellationToken)) WriteTag(builder, "N1", note, options.LineEnding, cancellationToken);
            WriteNativeFields(builder, item, BibliographyFormat.Ris, options.LineEnding, report, cancellationToken);
            WriteTag(builder, "ER", string.Empty, options.LineEnding, cancellationToken);
            if (itemIndex + 1 < document.Items.Count) builder.Append(options.LineEnding);
        }
        AddDocumentNativeLoss(document, BibliographyFormat.Ris, report, cancellationToken);
        return builder.ToString();
    }

    internal static string WriteNbib(BibliographyDocument document, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        var builder = new StringBuilder();
        string[] outputKeys = CodecMappings.OutputKeys(document.Items, BibliographyFormat.Nbib, cancellationToken);
        for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
            BibliographyItem item = document.Items[itemIndex];
            cancellationToken.ThrowIfCancellationRequested();
            bool hasNativePublicationTypes = HasNativeNbibPublicationTypes(item, cancellationToken);
            if (!hasNativePublicationTypes) WriteNbibPublicationTypes(builder, document.SourceFormat, item, options.LineEnding, report, cancellationToken);
            WriteTag(builder, "TI", item.Title, options.LineEnding, cancellationToken); WriteTag(builder, TaggedOutputTag(item, BibliographyFormat.Nbib, "container-title", "JT"), item.ContainerTitle, options.LineEnding, cancellationToken);
            foreach (BibliographyContributor author in Cancellable(item.Contributors, cancellationToken).Where(static contributor => contributor.Role == BibliographyContributorRole.Author && string.IsNullOrWhiteSpace(contributor.Name.Literal))) WriteTag(builder, "FAU", CodecMappings.FormatName(author.Name), options.LineEnding, cancellationToken);
            foreach (BibliographyContributor author in Cancellable(item.Contributors, cancellationToken).Where(static contributor => contributor.Role == BibliographyContributorRole.Author && string.IsNullOrWhiteSpace(contributor.Name.Literal))) WriteTag(builder, "AU", CompactName(author.Name), options.LineEnding, cancellationToken);
            foreach (BibliographyContributor author in Cancellable(item.Contributors, cancellationToken).Where(static contributor => contributor.Role == BibliographyContributorRole.Author && !string.IsNullOrWhiteSpace(contributor.Name.Literal))) WriteTag(builder, "CN", author.Name.Literal, options.LineEnding, cancellationToken);
            BibliographyDate? issued = Cancellable(item.Dates, cancellationToken).FirstOrDefault(static date => date.Role == BibliographyDateRole.Issued); if (issued != null) WriteTag(builder, "DP", CodecMappings.FormatDate(issued), options.LineEnding, cancellationToken);
            WriteTag(builder, "VI", item.Volume, options.LineEnding, cancellationToken); WriteTag(builder, "IP", item.Issue, options.LineEnding, cancellationToken); WriteTag(builder, "PG", item.Pages, options.LineEnding, cancellationToken);
            WriteTag(builder, "AB", item.Abstract, options.LineEnding, cancellationToken); WriteTag(builder, "LA", item.Language, options.LineEnding, cancellationToken);
            WriteNbibIdentifiers(builder, item, outputKeys[itemIndex], options.LineEnding, cancellationToken);
            foreach (string keyword in Cancellable(item.Keywords, cancellationToken)) WriteTag(builder, "OT", keyword, options.LineEnding, cancellationToken);
            foreach (string note in Cancellable(item.Notes, cancellationToken)) WriteTag(builder, "GN", note, options.LineEnding, cancellationToken);
            if (hasNativePublicationTypes) WriteNbibNativeFieldsAndPublicationTypes(builder, document.SourceFormat, item, options.LineEnding, report, cancellationToken);
            else WriteNativeFields(builder, item, BibliographyFormat.Nbib, options.LineEnding, report, cancellationToken);
            if (itemIndex + 1 < document.Items.Count) builder.Append(options.LineEnding);
        }
        AddDocumentNativeLoss(document, BibliographyFormat.Nbib, report, cancellationToken);
        return builder.ToString();
    }

    private static IList<BibliographyItem> Parse(string source, BibliographyFormat format, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, CancellationToken cancellationToken) {
        var items = new List<BibliographyItem>();
        var limits = new BibliographyLimitGuard(options);
        var diagnosticGuard = new BibliographyDiagnosticGuard(options, diagnostics, items);
        BibliographyItem? current = null;
        string? previousTag = null;
        BibliographyNativeField? previousNativeField = null;
        BibliographyIdentifier? previousIdentifier = null;
        bool previousIdentifierControlsKey = false;
        for (int offset = 0, lineIndex = 0; offset <= source.Length; lineIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            int lineOffset = offset;
            int lineEnd = FindLineEnd(source, lineOffset, cancellationToken);
            offset = lineEnd >= source.Length ? source.Length + 1 : lineEnd + (source[lineEnd] == '\r' && lineEnd + 1 < source.Length && source[lineEnd + 1] == '\n' ? 2 : 1);
            int contentStart = lineOffset;
            if (lineIndex == 0 && contentStart < lineEnd && source[contentStart] == '\uFEFF') contentStart++;
            if (IsWhiteSpace(source, contentStart, lineEnd, cancellationToken)) {
                if (format == BibliographyFormat.Nbib) current = null;
                previousTag = null; previousNativeField = null; previousIdentifier = null; previousIdentifierControlsKey = false; continue;
            }
            if (TrySplitLine(source, contentStart, lineEnd, items, limits, lineOffset, cancellationToken, out string tag, out string value)) {
                if (format == BibliographyFormat.Ris && string.Equals(tag, "TY", StringComparison.OrdinalIgnoreCase)) {
                    current = NewItem(items, limits, lineOffset); current.NativeType = value; current.Type = CodecMappings.ParseRisType(value);
                    previousTag = tag;
                    previousNativeField = null;
                    previousIdentifier = null;
                    previousIdentifierControlsKey = false;
                    continue;
                } else if (format == BibliographyFormat.Nbib && string.Equals(tag, "PMID", StringComparison.OrdinalIgnoreCase) && current != null && !string.IsNullOrEmpty(current.Key)) {
                    current = NewItem(items, limits, lineOffset);
                } else if (current == null) current = NewItem(items, limits, lineOffset);
                if (format == BibliographyFormat.Nbib && current!.Type == BibliographyItemType.Unknown) { current.Type = BibliographyItemType.ArticleJournal; current.NativeType = "Journal Article"; }
                if (format == BibliographyFormat.Ris && string.Equals(tag, "ER", StringComparison.OrdinalIgnoreCase)) {
                    if (value.Length > 0) diagnosticGuard.Add(new BibliographyDiagnostic("BIBTAG004", BibliographyDiagnosticSeverity.Warning, "RIS record terminator contains a value that cannot be preserved canonically.", offset: lineOffset, line: lineIndex + 1, column: 1, itemKey: current!.Key, field: tag));
                    current = null;
                    previousNativeField = null;
                    previousIdentifier = null;
                    previousIdentifierControlsKey = false;
                } else {
                    int nativeCount = current!.NativeFields.Count;
                    int identifierCount = current.Identifiers.Count;
                    bool accessionCanSetKey = format == BibliographyFormat.Ris && string.Equals(tag, "AN", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(current.Key);
                    Bind(current, format, tag, value);
                    previousNativeField = current.NativeFields.Count > nativeCount ? current.NativeFields[current.NativeFields.Count - 1] : null;
                    previousIdentifier = current.Identifiers.Count > identifierCount ? current.Identifiers[current.Identifiers.Count - 1] : null;
                    previousIdentifierControlsKey = accessionCanSetKey && previousIdentifier != null;
                }
                previousTag = tag;
            } else if (current != null && previousTag != null && (char.IsWhiteSpace(source[contentStart]) || format == BibliographyFormat.Ris)) {
                GetTrimmedRange(source, contentStart, lineEnd, out int continuationStart, out int continuationLength);
                limits.AddValue(items, continuationLength, lineOffset);
                string continuation = source.Substring(continuationStart, continuationLength);
                if (previousNativeField != null) {
                    previousNativeField.Value = AppendChecked(previousNativeField.Value, continuation, items, limits, lineOffset);
                    if (format == BibliographyFormat.Nbib && string.Equals(previousTag, "PT", StringComparison.OrdinalIgnoreCase)) UpdateNbibPublicationTypeAfterContinuation(current, previousNativeField, cancellationToken);
                }
                else AppendContinuation(current, format, previousTag, continuation, previousIdentifier, previousIdentifierControlsKey, diagnosticGuard, lineIndex + 1, lineOffset, items, limits);
            } else diagnosticGuard.Add(new BibliographyDiagnostic("BIBTAG001", BibliographyDiagnosticSeverity.Warning, $"Ignored malformed {format} line.", offset: lineOffset, line: lineIndex + 1, column: 1));
        }
        if (format == BibliographyFormat.Nbib) NormalizeNbibAuthors(items, cancellationToken);
        foreach (BibliographyItem item in items.Where(static item => string.IsNullOrWhiteSpace(item.Key))) diagnosticGuard.Add(new BibliographyDiagnostic("BIBTAG003", BibliographyDiagnosticSeverity.Warning, $"{format} record has no citation identifier.", itemKey: item.Key));
        return items;
    }

    private static BibliographyItem NewItem(IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) { limits.AddItem(items, offset); var item = new BibliographyItem(); items.Add(item); return item; }

    private static int FindLineEnd(string source, int start, CancellationToken cancellationToken) {
        int position = start;
        while (position < source.Length && source[position] != '\r' && source[position] != '\n') {
            if (((position - start) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            position++;
        }
        return position;
    }

    private static bool TrySplitLine(string source, int start, int end, IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset, CancellationToken cancellationToken, out string tag, out string value) {
        tag = string.Empty; value = string.Empty;
        if (start >= end || char.IsWhiteSpace(source[start])) return false;
        int dashPosition = start;
        while (dashPosition < end && source[dashPosition] != '-') {
            if (((dashPosition - start) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            dashPosition++;
        }
        if (dashPosition == start || dashPosition >= end) return false;
        int tagStart = start;
        int tagEnd = dashPosition;
        while (tagStart < tagEnd && char.IsWhiteSpace(source[tagStart])) tagStart++;
        while (tagEnd > tagStart && char.IsWhiteSpace(source[tagEnd - 1])) tagEnd--;
        int tagLength = tagEnd - tagStart;
        if (tagLength < 2 || tagLength > 5) return false;
        for (int index = tagStart; index < tagEnd; index++) if (!char.IsLetterOrDigit(source[index])) return false;
        int valueStart = dashPosition + 1;
        while (valueStart < end && char.IsWhiteSpace(source[valueStart])) valueStart++;
        int valueLength = end - valueStart;
        limits.AddValue(items, valueLength, offset);
        tag = source.Substring(tagStart, tagLength);
        value = source.Substring(valueStart, valueLength);
        return true;
    }

    private static bool IsWhiteSpace(string source, int start, int end, CancellationToken cancellationToken) {
        for (int index = start; index < end; index++) {
            if (((index - start) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (!char.IsWhiteSpace(source[index])) return false;
        }
        return true;
    }

    private static void GetTrimmedRange(string source, int start, int end, out int trimmedStart, out int length) {
        while (start < end && char.IsWhiteSpace(source[start])) start++;
        while (end > start && char.IsWhiteSpace(source[end - 1])) end--;
        trimmedStart = start;
        length = end - start;
    }

    private static void Bind(BibliographyItem item, BibliographyFormat format, string tag, string value) {
        string field = tag.ToUpperInvariant();
        if (format == BibliographyFormat.Ris) BindRis(item, field, tag, value); else BindNbib(item, field, tag, value);
    }

    private static void BindRis(BibliographyItem item, string field, string sourceTag, string value) {
        switch (field) {
            case "ID": SetScalar(item, BibliographyFormat.Ris, "key", sourceTag, value, assigned => item.Key = assigned); break;
            case "TI": case "T1": SetScalar(item, BibliographyFormat.Ris, "title", sourceTag, value, assigned => item.Title = assigned); break;
            case "T2": case "JF": case "JO": case "JA": SetScalar(item, BibliographyFormat.Ris, "container-title", sourceTag, value, assigned => item.ContainerTitle = assigned); break;
            case "AU": case "A1": AddTaggedContributor(item, BibliographyContributorRole.Author, sourceTag, value); break;
            case "ED": case "A2": AddTaggedContributor(item, BibliographyContributorRole.Editor, sourceTag, value); break;
            case "PY": case "Y1": case "DA": AddTaggedDate(item, BibliographyDateRole.Issued, sourceTag, value); break;
            case "Y2": AddTaggedDate(item, BibliographyDateRole.Accessed, sourceTag, value); break;
            case "PB": SetScalar(item, BibliographyFormat.Ris, "publisher", sourceTag, value, assigned => item.Publisher = assigned); break;
            case "CY": SetScalar(item, BibliographyFormat.Ris, "publisher-place", sourceTag, value, assigned => item.PublisherPlace = assigned); break;
            case "ET": SetScalar(item, BibliographyFormat.Ris, "edition", sourceTag, value, assigned => item.Edition = assigned); break;
            case "VL": SetScalar(item, BibliographyFormat.Ris, "volume", sourceTag, value, assigned => item.Volume = assigned); break;
            case "IS": SetScalar(item, BibliographyFormat.Ris, "issue", sourceTag, value, assigned => item.Issue = assigned); break;
            case "SP": SetPageStart(item, value); break;
            case "EP": SetPageEnd(item, value); break;
            case "AB": case "N2": SetScalar(item, BibliographyFormat.Ris, "abstract", sourceTag, value, assigned => item.Abstract = assigned); break;
            case "LA": SetScalar(item, BibliographyFormat.Ris, "language", sourceTag, value, assigned => item.Language = assigned); break;
            case "UR": case "L1": SetScalar(item, BibliographyFormat.Ris, "url", sourceTag, value, assigned => item.Url = assigned); break;
            case "DO": AddRisIdentifier(item, sourceTag, "DOI", value); break;
            case "SN": AddRisIdentifier(item, sourceTag, CodecMappings.InferSerialScheme(value), value); break;
            case "AN":
                if (string.IsNullOrWhiteSpace(item.Key)) { item.Key = value; item.TaggedScalarBindings.Add("Ris:key-from-accession"); }
                if (string.IsNullOrWhiteSpace(value)) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Ris, sourceTag, value));
                else ParseRisAccession(item, value);
                break;
            case "KW": item.Keywords.Add(value); break; case "N1": item.Notes.Add(value); break;
            default: item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Ris, sourceTag, value)); break;
        }
    }

    private static void BindNbib(BibliographyItem item, string field, string sourceTag, string value) {
        switch (field) {
            case "PMID": item.Key = value; AddTaggedIdentifier(item, "PMID", value, sourceTag); break;
            case "PT":
                var publicationType = new BibliographyNativeField(BibliographyFormat.Nbib, sourceTag, value);
                item.NativeFields.Add(publicationType);
                BindNbibPublicationType(item, publicationType);
                break;
            case "TI": SetScalar(item, BibliographyFormat.Nbib, "title", sourceTag, value, assigned => item.Title = assigned); break;
            case "JT": case "TA": SetScalar(item, BibliographyFormat.Nbib, "container-title", sourceTag, value, assigned => item.ContainerTitle = assigned); break;
            case "FAU": AddNbibContributor(item, sourceTag, CodecMappings.ParseCommaName(value)); break;
            case "AU": AddNbibContributor(item, sourceTag, ParseCompactNbibName(value)); break;
            case "CN": AddNbibContributor(item, sourceTag, new BibliographyName { Literal = value }); break;
            case "DP": AddTaggedDate(item, BibliographyDateRole.Issued, sourceTag, value); break;
            case "VI": SetScalar(item, BibliographyFormat.Nbib, "volume", sourceTag, value, assigned => item.Volume = assigned); break;
            case "IP": SetScalar(item, BibliographyFormat.Nbib, "issue", sourceTag, value, assigned => item.Issue = assigned); break;
            case "PG": SetScalar(item, BibliographyFormat.Nbib, "pages", sourceTag, value, assigned => item.Pages = assigned); break;
            case "AB": SetScalar(item, BibliographyFormat.Nbib, "abstract", sourceTag, value, assigned => item.Abstract = assigned); break;
            case "LA": SetScalar(item, BibliographyFormat.Nbib, "language", sourceTag, value, assigned => item.Language = assigned); break;
            case "LID": case "AID": ParseNbibIdentifier(item, value, sourceTag); break;
            case "IS":
                if (string.IsNullOrWhiteSpace(value)) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Nbib, sourceTag, value));
                else AddTaggedIdentifier(item, "ISSN", value, sourceTag);
                break;
            case "OT": item.Keywords.Add(value); break;
            case "MH": item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Nbib, sourceTag, value)); break;
            case "GN": item.Notes.Add(value); break;
            default: item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Nbib, sourceTag, value)); break;
        }
    }

    private static void ParseNbibIdentifier(BibliographyItem item, string value, string field) {
        int marker = value.LastIndexOf(" [", StringComparison.Ordinal);
        if (marker > 0 && value.EndsWith("]", StringComparison.Ordinal)) {
            string scheme = value.Substring(marker + 2, value.Length - marker - 3).Trim();
            string identifierValue = value.Substring(0, marker).Trim();
            if (!string.IsNullOrWhiteSpace(scheme) && !string.IsNullOrWhiteSpace(identifierValue)) AddTaggedIdentifier(item, scheme, identifierValue, field);
            else item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Nbib, field, value));
        }
        else item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Nbib, field, value));
    }

    private static void BindNbibPublicationType(BibliographyItem item, BibliographyNativeField field) {
        BibliographyItemType parsed = CodecMappings.ParseType(field.Value);
        if (parsed == BibliographyItemType.Unknown || item.NbibTypeBinding != null) return;
        item.Type = parsed;
        item.NativeType = field.Value;
        item.NbibTypeBinding = field;
        item.TaggedScalarBindings.Add("Nbib:type");
    }

    private static void UpdateNbibPublicationTypeAfterContinuation(BibliographyItem item, BibliographyNativeField field, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (ReferenceEquals(item.NbibTypeBinding, field)) {
            BibliographyItemType parsed = CodecMappings.ParseType(field.Value);
            if (parsed != BibliographyItemType.Unknown) {
                item.Type = parsed;
                item.NativeType = field.Value;
                return;
            }
            item.Type = BibliographyItemType.ArticleJournal;
            item.NativeType = "Journal Article";
            item.NbibTypeBinding = null;
            item.TaggedScalarBindings.Remove("Nbib:type");
        }
        if (item.NbibTypeBinding == null) BindNbibPublicationType(item, field);
    }

    private static void AddTaggedIdentifier(BibliographyItem item, string scheme, string value, string tag) {
        if (string.IsNullOrWhiteSpace(value)) return;
        var identifier = new BibliographyIdentifier(scheme, value);
        item.Identifiers.Add(identifier);
        item.TaggedIdentifierTags[identifier] = tag;
    }

    private static void ParseRisAccession(BibliographyItem item, string value) {
        int separator = value.IndexOf(':');
        if (separator > 0 && separator + 1 < value.Length) CodecMappings.AddIdentifier(item, value.Substring(0, separator), value.Substring(separator + 1));
        else CodecMappings.AddIdentifier(item, "accession", value);
    }

    private static void AddRisIdentifier(BibliographyItem item, string tag, string scheme, string value) {
        if (string.IsNullOrWhiteSpace(value)) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Ris, tag, value));
        else CodecMappings.AddIdentifier(item, scheme, value);
    }

    private static void SetScalar(BibliographyItem item, BibliographyFormat format, string semanticName, string sourceTag, string value, Action<string> write) {
        string binding = format + ":" + semanticName;
        if (item.TaggedScalarBindings.Add(binding)) { write(value); item.TaggedFieldNames[binding] = sourceTag; }
        else item.NativeFields.Add(new BibliographyNativeField(format, sourceTag, value));
    }

    private static void AddTaggedContributor(BibliographyItem item, BibliographyContributorRole role, string tag, string value) {
        var contributor = new BibliographyContributor(role, CodecMappings.ParseCommaName(value));
        item.Contributors.Add(contributor);
        item.TaggedContributorTags[contributor] = tag;
    }

    private static void AddTaggedDate(BibliographyItem item, BibliographyDateRole role, string tag, string value) {
        BibliographyDate date = CodecMappings.ParseDate(role, value);
        item.Dates.Add(date);
        item.TaggedDateTags[date] = tag;
    }

    private static void SetPageEnd(BibliographyItem item, string value) {
        const string binding = "Ris:pages-end";
        if (item.TaggedScalarBindings.Add(binding)) { item.RisPageEnd = value; UpdateRisPages(item); }
        else item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Ris, "EP", value));
    }

    private static void SetPageStart(BibliographyItem item, string value) {
        const string binding = "Ris:pages-start";
        if (item.TaggedScalarBindings.Add(binding)) { item.RisPageStart = value; UpdateRisPages(item); }
        else item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Ris, "SP", value));
    }

    private static void UpdateRisPages(BibliographyItem item) {
        bool hasStart = item.TaggedScalarBindings.Contains("Ris:pages-start");
        bool hasEnd = item.TaggedScalarBindings.Contains("Ris:pages-end");
        item.Pages = hasStart && hasEnd ? item.RisPageStart + "-" + item.RisPageEnd : hasStart ? item.RisPageStart : hasEnd ? item.RisPageEnd : null;
    }

    private static void AppendContinuation(BibliographyItem item, BibliographyFormat format, string tag, string value, BibliographyIdentifier? identifierBinding, bool identifierControlsKey, BibliographyDiagnosticGuard diagnostics, int line, int offset, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        string field = tag.ToUpperInvariant();
        if (format == BibliographyFormat.Ris) {
            switch (field) {
                case "TI": case "T1": item.Title = AppendChecked(item.Title, value, items, limits, offset); return;
                case "T2": case "JF": case "JO": case "JA": item.ContainerTitle = AppendChecked(item.ContainerTitle, value, items, limits, offset); return;
                case "ID": item.Key = AppendChecked(item.Key, value, items, limits, offset); return;
                case "PB": item.Publisher = AppendChecked(item.Publisher, value, items, limits, offset); return;
                case "CY": item.PublisherPlace = AppendChecked(item.PublisherPlace, value, items, limits, offset); return;
                case "ET": item.Edition = AppendChecked(item.Edition, value, items, limits, offset); return;
                case "VL": item.Volume = AppendChecked(item.Volume, value, items, limits, offset); return;
                case "IS": item.Issue = AppendChecked(item.Issue, value, items, limits, offset); return;
                case "SP": item.RisPageStart = AppendChecked(item.RisPageStart, value, items, limits, offset); UpdateRisPages(item); return;
                case "EP": item.RisPageEnd = AppendChecked(item.RisPageEnd, value, items, limits, offset); UpdateRisPages(item); return;
                case "AB": case "N2": item.Abstract = AppendChecked(item.Abstract, value, items, limits, offset); return;
                case "LA": item.Language = AppendChecked(item.Language, value, items, limits, offset); return;
                case "UR": case "L1": item.Url = AppendChecked(item.Url, value, items, limits, offset); return;
                case "PY": case "Y1": case "DA": AppendDate(item, BibliographyDateRole.Issued, field, value, items, limits, offset); return;
                case "Y2": AppendDate(item, BibliographyDateRole.Accessed, field, value, items, limits, offset); return;
                case "DO": AppendIdentifier(identifierBinding, value, items, limits, offset); return;
                case "SN": AppendIdentifier(identifierBinding, value, items, limits, offset); return;
                case "AN": if (identifierControlsKey) item.Key = AppendChecked(item.Key, value, items, limits, offset); AppendIdentifier(identifierBinding, value, items, limits, offset); return;
                case "N1": AppendLast(item.Notes, value, items, limits, offset); return;
                case "KW": AppendLast(item.Keywords, value, items, limits, offset); return;
                case "AU": case "A1": AppendContributor(item, BibliographyContributorRole.Author, value, items, limits, offset); return;
                case "ED": case "A2": AppendContributor(item, BibliographyContributorRole.Editor, value, items, limits, offset); return;
            }
        } else {
            switch (field) {
                case "TI": item.Title = AppendChecked(item.Title, value, items, limits, offset); return;
                case "JT": case "TA": item.ContainerTitle = AppendChecked(item.ContainerTitle, value, items, limits, offset); return;
                case "PMID": item.Key = AppendChecked(item.Key, value, items, limits, offset); AppendTaggedIdentifier(item, field, value, items, limits, offset); return;
                case "DP": AppendDate(item, BibliographyDateRole.Issued, field, value, items, limits, offset); return;
                case "VI": item.Volume = AppendChecked(item.Volume, value, items, limits, offset); return;
                case "IP": item.Issue = AppendChecked(item.Issue, value, items, limits, offset); return;
                case "PG": item.Pages = AppendChecked(item.Pages, value, items, limits, offset); return;
                case "AB": item.Abstract = AppendChecked(item.Abstract, value, items, limits, offset); return;
                case "LA": item.Language = AppendChecked(item.Language, value, items, limits, offset); return;
                case "IS": case "LID": case "AID": AppendTaggedIdentifier(item, field, value, items, limits, offset); return;
                case "GN": AppendLast(item.Notes, value, items, limits, offset); return;
                case "OT": AppendLast(item.Keywords, value, items, limits, offset); return;
                case "FAU": case "AU": case "CN": AppendNbibContributor(item, field, value, items, limits, offset); return;
            }
        }
        if (item.NativeFields.LastOrDefault(nativeField => nativeField.Format == format && string.Equals(nativeField.Name, tag, StringComparison.OrdinalIgnoreCase)) is BibliographyNativeField native) {
            native.Value = AppendChecked(native.Value, value, items, limits, offset);
        } else {
            item.NativeFields.Add(new BibliographyNativeField(format, tag, value));
            diagnostics.Add(new BibliographyDiagnostic("BIBTAG002", BibliographyDiagnosticSeverity.Information, $"Continuation for '{tag}' was retained as a native field.", offset: offset, line: line, column: 1, itemKey: item.Key, field: tag));
        }
    }

    private static string Append(string? current, string value) => string.IsNullOrEmpty(current) ? value : current + " " + value;
    private static string AppendChecked(string? current, string value, IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) {
        int currentLength = current?.Length ?? 0;
        int separatorLength = string.IsNullOrEmpty(current) ? 0 : 1;
        limits.CheckAdditionalValueLength(items, currentLength, checked(separatorLength + value.Length), offset);
        return Append(current, value);
    }
    private static void AppendLast(IList<string> values, string continuation, IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) { if (values.Count == 0) values.Add(continuation); else values[values.Count - 1] = AppendChecked(values[values.Count - 1], continuation, items, limits, offset); }
    private static void AppendContributor(BibliographyItem item, BibliographyContributorRole role, string continuation, IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) {
        BibliographyContributor? contributor = item.Contributors.LastOrDefault(value => value.Role == role);
        if (contributor == null) { item.Contributors.Add(new BibliographyContributor(role, CodecMappings.ParseCommaName(continuation))); return; }
        if (!string.IsNullOrWhiteSpace(contributor.Name.Literal)) { contributor.Name.Literal = AppendChecked(contributor.Name.Literal, continuation, items, limits, offset); return; }
        string combined = AppendChecked(CodecMappings.FormatName(contributor.Name), continuation, items, limits, offset);
        BibliographyName parsed = CodecMappings.ParseCommaName(combined);
        contributor.Name.Given = parsed.Given; contributor.Name.Family = parsed.Family; contributor.Name.Literal = parsed.Literal; contributor.Name.Suffix = parsed.Suffix;
        contributor.Name.DroppingParticle = parsed.DroppingParticle; contributor.Name.NonDroppingParticle = parsed.NonDroppingParticle;
    }

    private static void AddNbibContributor(BibliographyItem item, string tag, BibliographyName name) {
        var contributor = new BibliographyContributor(BibliographyContributorRole.Author, name);
        item.Contributors.Add(contributor);
        item.TaggedContributorTags[contributor] = tag;
    }

    private static void AppendNbibContributor(BibliographyItem item, string tag, string continuation, IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) {
        BibliographyContributor? contributor = item.Contributors.LastOrDefault(candidate => item.TaggedContributorTags.TryGetValue(candidate, out string? sourceTag) && string.Equals(sourceTag, tag, StringComparison.OrdinalIgnoreCase));
        if (contributor == null) return;
        string original = string.Equals(tag, "AU", StringComparison.OrdinalIgnoreCase) ? CompactName(contributor.Name) : string.Equals(tag, "CN", StringComparison.OrdinalIgnoreCase) ? contributor.Name.Literal ?? string.Empty : CodecMappings.FormatName(contributor.Name);
        BibliographyName parsed = string.Equals(tag, "AU", StringComparison.OrdinalIgnoreCase) ? ParseCompactNbibName(AppendChecked(original, continuation, items, limits, offset)) : string.Equals(tag, "CN", StringComparison.OrdinalIgnoreCase) ? new BibliographyName { Literal = AppendChecked(original, continuation, items, limits, offset) } : CodecMappings.ParseCommaName(AppendChecked(original, continuation, items, limits, offset));
        contributor.Name.Given = parsed.Given; contributor.Name.Family = parsed.Family; contributor.Name.Literal = parsed.Literal; contributor.Name.Suffix = parsed.Suffix;
        contributor.Name.DroppingParticle = parsed.DroppingParticle; contributor.Name.NonDroppingParticle = parsed.NonDroppingParticle;
    }

    private static void AppendDate(BibliographyItem item, BibliographyDateRole role, string tag, string continuation, IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) {
        BibliographyDate? date = item.Dates.LastOrDefault(candidate => candidate.Role == role && item.TaggedDateTags.TryGetValue(candidate, out string? sourceTag) && string.Equals(sourceTag, tag, StringComparison.OrdinalIgnoreCase));
        if (date == null) return;
        string combined = AppendChecked(CodecMappings.FormatDate(date), continuation, items, limits, offset);
        BibliographyDate parsed = CodecMappings.ParseDate(role, combined);
        date.Year = parsed.Year; date.Month = parsed.Month; date.Day = parsed.Day; date.EndYear = parsed.EndYear; date.EndMonth = parsed.EndMonth; date.EndDay = parsed.EndDay; date.Literal = parsed.Literal;
    }

    private static void AppendIdentifier(BibliographyIdentifier? identifier, string continuation, IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) {
        if (identifier != null) identifier.Value = AppendChecked(identifier.Value, continuation, items, limits, offset);
    }

    private static void AppendTaggedIdentifier(BibliographyItem item, string tag, string continuation, IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) {
        BibliographyIdentifier? identifier = item.Identifiers.LastOrDefault(candidate => item.TaggedIdentifierTags.TryGetValue(candidate, out string? sourceTag) && string.Equals(sourceTag, tag, StringComparison.OrdinalIgnoreCase));
        if (identifier != null) identifier.Value = AppendChecked(identifier.Value, continuation, items, limits, offset);
    }

    internal static void NormalizeNbibAuthors(IEnumerable<BibliographyItem> items, CancellationToken cancellationToken) {
        foreach (BibliographyItem item in items) {
            cancellationToken.ThrowIfCancellationRequested();
            var compactAuthors = new List<BibliographyContributor>();
            var fullAuthorsByCompactName = new Dictionary<string, Queue<BibliographyContributor>>(StringComparer.OrdinalIgnoreCase);
            var contributorIndexes = new Dictionary<BibliographyContributor, int>();
            for (int index = 0; index < item.Contributors.Count; index++) {
                if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                BibliographyContributor contributor = item.Contributors[index];
                contributorIndexes[contributor] = index;
                if (!item.TaggedContributorTags.TryGetValue(contributor, out string? tag)) continue;
                if (string.Equals(tag, "AU", StringComparison.OrdinalIgnoreCase)) compactAuthors.Add(contributor);
                else if (string.Equals(tag, "FAU", StringComparison.OrdinalIgnoreCase)) {
                    string key = NormalizeCompactName(CompactName(contributor.Name), cancellationToken);
                    if (!fullAuthorsByCompactName.TryGetValue(key, out Queue<BibliographyContributor>? matches)) {
                        matches = new Queue<BibliographyContributor>();
                        fullAuthorsByCompactName.Add(key, matches);
                    }
                    matches.Enqueue(contributor);
                }
            }

            var removed = new HashSet<BibliographyContributor>();
            var replacements = new Dictionary<BibliographyContributor, BibliographyContributor>();
            foreach (BibliographyContributor compact in compactAuthors) {
                cancellationToken.ThrowIfCancellationRequested();
                string key = NormalizeCompactName(ParsedNbibCompactName(compact.Name), cancellationToken);
                if (!fullAuthorsByCompactName.TryGetValue(key, out Queue<BibliographyContributor>? matches) || matches.Count == 0) continue;
                BibliographyContributor full = matches.Dequeue();
                if (contributorIndexes[compact] < contributorIndexes[full]) {
                    replacements.Add(compact, full);
                    removed.Add(full);
                } else removed.Add(compact);
                item.TaggedContributorTags.Remove(compact);
            }

            if (removed.Count == 0) continue;
            var normalized = new List<BibliographyContributor>(item.Contributors.Count - removed.Count);
            foreach (BibliographyContributor contributor in item.Contributors) {
                cancellationToken.ThrowIfCancellationRequested();
                if (replacements.TryGetValue(contributor, out BibliographyContributor? replacement)) normalized.Add(replacement);
                else if (!removed.Contains(contributor)) normalized.Add(contributor);
            }
            item.Contributors.Clear();
            foreach (BibliographyContributor contributor in normalized) item.Contributors.Add(contributor);
        }
    }

    private static string NormalizeCompactName(string value, CancellationToken cancellationToken) {
        var builder = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (!char.IsLetterOrDigit(value, index)) continue;
            builder.Append(value[index]);
            if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) builder.Append(value[++index]);
        }
        return builder.ToString();
    }
    private static string ParsedNbibCompactName(BibliographyName name) =>
        string.IsNullOrWhiteSpace(name.Literal) ? ((name.Family ?? string.Empty) + " " + (name.Given ?? string.Empty)).Trim() : name.Literal!;
    private static void WriteTag(StringBuilder builder, string tag, string? value, string lineEnding, CancellationToken cancellationToken) {
        if (value == null) return;
        cancellationToken.ThrowIfCancellationRequested();
        builder.Append(tag);
        for (int index = tag.Length; index < 4; index++) builder.Append(' ');
        builder.Append("- ");
        int segmentStart = 0;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char current = value[index];
            if (current != '\r' && current != '\n') continue;
            AppendCancellable(builder, value, segmentStart, index - segmentStart, cancellationToken);
            builder.Append(lineEnding).Append("      ");
            if (current == '\r' && index + 1 < value.Length && value[index + 1] == '\n') index++;
            segmentStart = index + 1;
        }
        AppendCancellable(builder, value, segmentStart, value.Length - segmentStart, cancellationToken);
        builder.Append(lineEnding);
        cancellationToken.ThrowIfCancellationRequested();
    }
    private static void AppendCancellable(StringBuilder builder, string value, int start, int length, CancellationToken cancellationToken) {
        int end = checked(start + length);
        while (start < end) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(4096, end - start);
            builder.Append(value, start, count);
            start += count;
        }
    }
    private static void WriteRisPages(StringBuilder builder, BibliographyItem item, string lineEnding, CancellationToken cancellationToken) {
        GetRisPageOutput(item, out bool writeStart, out string? start, out bool writeEnd, out string? end);
        if (writeStart) WriteTag(builder, "SP", start, lineEnding, cancellationToken);
        if (writeEnd) WriteTag(builder, "EP", end, lineEnding, cancellationToken);
    }

    private static void GetRisPageOutput(BibliographyItem item, out bool writeStart, out string? start, out bool writeEnd, out string? end) {
        bool sourceHasStart = item.TaggedScalarBindings.Contains("Ris:pages-start");
        bool sourceHasEnd = item.TaggedScalarBindings.Contains("Ris:pages-end");
        string? sourcePages = sourceHasStart && sourceHasEnd ? item.RisPageStart + "-" + item.RisPageEnd : sourceHasStart ? item.RisPageStart : sourceHasEnd ? item.RisPageEnd : null;
        if ((sourceHasStart || sourceHasEnd) && string.Equals(item.Pages, sourcePages, StringComparison.Ordinal)) {
            writeStart = sourceHasStart; start = item.RisPageStart;
            writeEnd = sourceHasEnd; end = item.RisPageEnd;
            return;
        }
        if (item.Pages == null) {
            writeStart = false; start = null; writeEnd = false; end = null;
            return;
        }
        string[] parts = item.Pages.Split(new[] { '-' }, 2);
        writeStart = true; start = parts[0];
        writeEnd = parts.Length > 1; end = writeEnd ? parts[1] : null;
    }
    private static void WriteDateTags(StringBuilder builder, BibliographyItem item, string lineEnding, string issuedTag, string accessedTag, CancellationToken cancellationToken) {
        bool wroteIssued = false;
        bool wroteAccessed = false;
        foreach (BibliographyDate date in Cancellable(item.Dates, cancellationToken)) {
            if (date.Role == BibliographyDateRole.Issued && !wroteIssued) {
                WriteTag(builder, DateTag(item, date, issuedTag, "Y1", "DA"), CodecMappings.FormatDate(date), lineEnding, cancellationToken);
                wroteIssued = true;
            } else if (date.Role == BibliographyDateRole.Accessed && !wroteAccessed) {
                WriteTag(builder, DateTag(item, date, accessedTag, "Y2"), CodecMappings.FormatDate(date), lineEnding, cancellationToken);
                wroteAccessed = true;
            }
        }
    }
    private static void WriteRisIdentifier(StringBuilder builder, BibliographyIdentifier identifier, string lineEnding, CancellationToken cancellationToken) { if (string.Equals(identifier.Scheme, "DOI", StringComparison.OrdinalIgnoreCase)) WriteTag(builder, "DO", identifier.Value, lineEnding, cancellationToken); else if (string.Equals(identifier.Scheme, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "SN", StringComparison.OrdinalIgnoreCase)) WriteTag(builder, "SN", identifier.Value, lineEnding, cancellationToken); else if (string.Equals(identifier.Scheme, "accession", StringComparison.OrdinalIgnoreCase)) WriteTag(builder, "AN", identifier.Value.IndexOf(':') >= 0 ? "accession:" + identifier.Value : identifier.Value, lineEnding, cancellationToken); else WriteTag(builder, "AN", identifier.Scheme + ":" + identifier.Value, lineEnding, cancellationToken); }
    internal static bool CanRoundTripRisIdentifier(BibliographyIdentifier identifier) {
        if (string.Equals(identifier.Scheme, "DOI", StringComparison.OrdinalIgnoreCase)) return string.Equals(identifier.Scheme, "DOI", StringComparison.Ordinal);
        if (string.Equals(identifier.Scheme, "accession", StringComparison.OrdinalIgnoreCase)) return string.Equals(identifier.Scheme, "accession", StringComparison.Ordinal);
        if (string.Equals(identifier.Scheme, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "SN", StringComparison.OrdinalIgnoreCase)) return string.Equals(identifier.Scheme, CodecMappings.InferSerialScheme(identifier.Value), StringComparison.Ordinal);
        return !string.IsNullOrWhiteSpace(identifier.Scheme) && identifier.Scheme.IndexOf(':') < 0 && identifier.Scheme.IndexOf('\r') < 0 && identifier.Scheme.IndexOf('\n') < 0;
    }
    internal static bool ProtectsLeadingWhitespaceInRisIdentifier(BibliographyIdentifier identifier) =>
        CanRoundTripRisIdentifier(identifier) &&
        !string.Equals(identifier.Scheme, "DOI", StringComparison.OrdinalIgnoreCase) &&
        !string.Equals(identifier.Scheme, "accession", StringComparison.OrdinalIgnoreCase) &&
        !string.Equals(identifier.Scheme, "ISBN", StringComparison.OrdinalIgnoreCase) &&
        !string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase) &&
        !string.Equals(identifier.Scheme, "SN", StringComparison.OrdinalIgnoreCase);
    internal static bool CanRoundTripRisType(BibliographyItemType type) =>
        type != BibliographyItemType.Unknown && CodecMappings.ParseRisType(CodecMappings.ToRisType(type)) == type;

    internal static bool CanPreserveNativeType(BibliographyFormat sourceFormat, BibliographyItem item) =>
        sourceFormat == BibliographyFormat.Ris && IsRisType(item.NativeType) && CodecMappings.ParseRisType(item.NativeType) == item.Type;

    internal static bool CanPreserveUnknownRisType(string? nativeType) =>
        IsRisType(nativeType) && CodecMappings.ParseRisType(nativeType) == BibliographyItemType.Unknown;

    private static bool HasNativeNbibPublicationTypes(BibliographyItem item, CancellationToken cancellationToken) =>
        Cancellable(item.NativeFields, cancellationToken).Any(field => field.Format == BibliographyFormat.Nbib && string.Equals(field.Name, "PT", StringComparison.OrdinalIgnoreCase));

    private static void WriteNbibPublicationTypes(StringBuilder builder, BibliographyFormat sourceFormat, BibliographyItem item, string lineEnding, BibliographyConversionReport report, CancellationToken cancellationToken) {
        BibliographyNativeField[] nativeTypes = Cancellable(item.NativeFields, cancellationToken).Where(field => field.Format == BibliographyFormat.Nbib && string.Equals(field.Name, "PT", StringComparison.OrdinalIgnoreCase)).ToArray();
        BibliographyItemType sourceType = nativeTypes.Select(field => CodecMappings.ParseType(field.Value)).FirstOrDefault(static type => type != BibliographyItemType.Unknown);
        bool preserveRecognizedSourceTypes = sourceType == item.Type;
        bool wroteTypedValue = false;
        foreach (BibliographyNativeField field in Cancellable(nativeTypes, cancellationToken)) {
            wroteTypedValue |= WriteNbibPublicationTypeField(builder, item, field, preserveRecognizedSourceTypes, lineEnding, report, cancellationToken);
        }
        WriteNbibFallbackPublicationType(builder, sourceFormat, item, sourceType, wroteTypedValue, lineEnding, cancellationToken);
    }

    private static void WriteNbibNativeFieldsAndPublicationTypes(StringBuilder builder, BibliographyFormat sourceFormat, BibliographyItem item, string lineEnding, BibliographyConversionReport report, CancellationToken cancellationToken) {
        BibliographyNativeField[] nativeTypes = Cancellable(item.NativeFields, cancellationToken).Where(field => field.Format == BibliographyFormat.Nbib && string.Equals(field.Name, "PT", StringComparison.OrdinalIgnoreCase)).ToArray();
        BibliographyItemType sourceType = nativeTypes.Select(field => CodecMappings.ParseType(field.Value)).FirstOrDefault(static type => type != BibliographyItemType.Unknown);
        bool preserveRecognizedSourceTypes = sourceType == item.Type;
        bool wroteTypedValue = false;
        foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken)) {
            if (field.Format == BibliographyFormat.Nbib && string.Equals(field.Name, "PT", StringComparison.OrdinalIgnoreCase)) {
                wroteTypedValue |= WriteNbibPublicationTypeField(builder, item, field, preserveRecognizedSourceTypes, lineEnding, report, cancellationToken);
            } else WriteNativeField(builder, item, field, BibliographyFormat.Nbib, lineEnding, report, cancellationToken);
        }
        WriteNbibFallbackPublicationType(builder, sourceFormat, item, sourceType, wroteTypedValue, lineEnding, cancellationToken);
    }

    private static bool WriteNbibPublicationTypeField(StringBuilder builder, BibliographyItem item, BibliographyNativeField field, bool preserveRecognizedSourceTypes, string lineEnding, BibliographyConversionReport report, CancellationToken cancellationToken) {
        BibliographyItemType parsed = CodecMappings.ParseType(field.Value);
        if (parsed == BibliographyItemType.Unknown) {
            WriteTag(builder, field.Name, field.Value, lineEnding, cancellationToken);
            report.Add("BIBCONV013", BibliographyDiagnosticSeverity.Information, "Preserved an unrecognized NBIB publication type.", BibliographyConversionAction.PreservedExtension, item, field.Name);
            return false;
        }
        if (preserveRecognizedSourceTypes || parsed == item.Type) {
            WriteTag(builder, field.Name, field.Value, lineEnding, cancellationToken);
            return parsed == item.Type;
        }
        report.Add("BIBCONV122", BibliographyDiagnosticSeverity.Warning, $"Recognized NBIB publication type '{field.Value}' conflicts with the edited typed item kind and was omitted.", BibliographyConversionAction.Omitted, item, field.Name);
        return false;
    }

    private static void WriteNbibFallbackPublicationType(StringBuilder builder, BibliographyFormat sourceFormat, BibliographyItem item, BibliographyItemType sourceType, bool wroteTypedValue, string lineEnding, CancellationToken cancellationToken) {
        bool preserveSourceAbsence = sourceFormat == BibliographyFormat.Nbib && item.Type == BibliographyItemType.ArticleJournal &&
            string.Equals(item.NativeType, "Journal Article", StringComparison.Ordinal) && sourceType == BibliographyItemType.Unknown;
        if (!wroteTypedValue && !preserveSourceAbsence && TryGetNbibPublicationType(item.Type, out string? publicationType)) WriteTag(builder, "PT", publicationType, lineEnding, cancellationToken);
    }

    private static void WriteNbibIdentifiers(StringBuilder builder, BibliographyItem item, string fallbackKey, string lineEnding, CancellationToken cancellationToken) {
        bool wrotePmid = false;
        foreach (BibliographyIdentifier identifier in Cancellable(item.Identifiers, cancellationToken)) {
            if (string.Equals(identifier.Scheme, "PMID", StringComparison.OrdinalIgnoreCase)) {
                if (!wrotePmid) WriteTag(builder, "PMID", identifier.Value, lineEnding, cancellationToken);
                wrotePmid = true;
            } else WriteNbibIdentifier(builder, item, identifier, lineEnding, cancellationToken);
        }
        if (!wrotePmid) WriteTag(builder, "PMID", fallbackKey, lineEnding, cancellationToken);
    }

    private static void WriteNbibIdentifier(StringBuilder builder, BibliographyItem item, BibliographyIdentifier identifier, string lineEnding, CancellationToken cancellationToken) {
        if (string.Equals(identifier.Scheme, "PMID", StringComparison.OrdinalIgnoreCase)) return;
        if (string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase)) WriteTag(builder, "IS", identifier.Value, lineEnding, cancellationToken);
        else if (CanRoundTripNbibIdentifier(identifier)) WriteTag(builder, NbibIdentifierTag(item, identifier), identifier.Value + " [" + identifier.Scheme + "]", lineEnding, cancellationToken);
    }
    internal static bool CanRoundTripNbibIdentifier(BibliographyIdentifier identifier) =>
        string.Equals(identifier.Scheme, "PMID", StringComparison.Ordinal) ||
        string.Equals(identifier.Scheme, "ISSN", StringComparison.Ordinal) ||
        !string.Equals(identifier.Scheme, "PMID", StringComparison.OrdinalIgnoreCase) && !string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase) &&
        string.Equals(identifier.Value, identifier.Value.Trim(), StringComparison.Ordinal) && identifier.Scheme.IndexOf(" [", StringComparison.Ordinal) < 0;
    internal static bool CanRoundTripNbibType(BibliographyItemType type) => TryGetNbibPublicationType(type, out _);
    private static bool TryGetNbibPublicationType(BibliographyItemType type, out string? value) {
        switch (type) {
            case BibliographyItemType.ArticleJournal: value = "Journal Article"; return true;
            case BibliographyItemType.ArticleMagazine: value = "Magazine Article"; return true;
            case BibliographyItemType.ArticleNewspaper: value = "Newspaper Article"; return true;
            case BibliographyItemType.Book: value = "Book"; return true;
            case BibliographyItemType.Chapter: value = "Book Chapter"; return true;
            case BibliographyItemType.PaperConference: value = "Conference Paper"; return true;
            case BibliographyItemType.Proceedings: value = "Proceedings"; return true;
            case BibliographyItemType.Report: value = "Report"; return true;
            case BibliographyItemType.Thesis: value = "Thesis"; return true;
            case BibliographyItemType.WebPage: value = "Web Page"; return true;
            case BibliographyItemType.Dataset: value = "Dataset"; return true;
            case BibliographyItemType.Software: value = "Computer Program"; return true;
            case BibliographyItemType.Patent: value = "Patent"; return true;
            case BibliographyItemType.LegalCase: value = "legal_case"; return true;
            case BibliographyItemType.Manuscript: value = "Manuscript"; return true;
            case BibliographyItemType.PersonalCommunication: value = "personal_communication"; return true;
            case BibliographyItemType.Document: value = "Document"; return true;
            default: value = null; return false;
        }
    }
    private static string NbibIdentifierTag(BibliographyItem item, BibliographyIdentifier identifier) => item.TaggedIdentifierTags.TryGetValue(identifier, out string? sourceTag) && (string.Equals(sourceTag, "LID", StringComparison.OrdinalIgnoreCase) || string.Equals(sourceTag, "AID", StringComparison.OrdinalIgnoreCase)) ? sourceTag.ToUpperInvariant() : "AID";
    private static BibliographyName ParseCompactNbibName(string value) {
        string trimmed = value.Trim();
        if (trimmed.IndexOf(',') >= 0) return CodecMappings.ParseCommaName(trimmed);
        string[] words = trimmed.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0) return new BibliographyName();
        if (words.Length == 1) return new BibliographyName { Family = words[0] };
        return new BibliographyName { Family = string.Join(" ", words.Take(words.Length - 1)), Given = words[words.Length - 1] };
    }
    private static string TaggedOutputTag(BibliographyItem item, BibliographyFormat format, string semanticName, string fallback) =>
        item.TaggedFieldNames.TryGetValue(format + ":" + semanticName, out string? sourceTag) && IsTag(sourceTag) ? sourceTag.ToUpperInvariant() : fallback;
    private static string ContributorTag(BibliographyItem item, BibliographyContributor contributor, string fallback, params string[] alternatives) =>
        item.TaggedContributorTags.TryGetValue(contributor, out string? sourceTag) && (string.Equals(sourceTag, fallback, StringComparison.OrdinalIgnoreCase) || alternatives.Any(alternative => string.Equals(sourceTag, alternative, StringComparison.OrdinalIgnoreCase))) ? sourceTag.ToUpperInvariant() : fallback;
    private static string DateTag(BibliographyItem item, BibliographyDate date, string fallback, params string[] alternatives) =>
        item.TaggedDateTags.TryGetValue(date, out string? sourceTag) && (string.Equals(sourceTag, fallback, StringComparison.OrdinalIgnoreCase) || alternatives.Any(alternative => string.Equals(sourceTag, alternative, StringComparison.OrdinalIgnoreCase))) ? sourceTag.ToUpperInvariant() : fallback;
    private static string CompactName(BibliographyName name) {
        if (!string.IsNullOrWhiteSpace(name.Literal)) return name.Literal!;
        string initials = Initials(name.Given);
        if (initials.Length == 0 && name.Family?.Any(char.IsWhiteSpace) == true) return CodecMappings.FormatName(name);
        return ((name.Family ?? string.Empty) + " " + initials).Trim();
    }
    private static string Initials(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var builder = new StringBuilder();
        foreach (string part in value!.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries)) builder.Append(StringInfo.GetNextTextElement(part, 0));
        return builder.ToString();
    }

    private static void WriteNativeFields(StringBuilder builder, BibliographyItem item, BibliographyFormat format, string lineEnding, BibliographyConversionReport report, CancellationToken cancellationToken) {
        foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken)) {
            if (format == BibliographyFormat.Nbib && field.Format == format && string.Equals(field.Name, "PT", StringComparison.OrdinalIgnoreCase)) continue;
            WriteNativeField(builder, item, field, format, lineEnding, report, cancellationToken);
        }
    }

    private static void WriteNativeField(StringBuilder builder, BibliographyItem item, BibliographyNativeField field, BibliographyFormat format, string lineEnding, BibliographyConversionReport report, CancellationToken cancellationToken) {
        bool unsafeBoundary = format == BibliographyFormat.Ris && (string.Equals(field.Name, "TY", StringComparison.OrdinalIgnoreCase) || string.Equals(field.Name, "ER", StringComparison.OrdinalIgnoreCase)) || format == BibliographyFormat.Nbib && string.Equals(field.Name, "PMID", StringComparison.OrdinalIgnoreCase);
        if (field.Format == format && IsTag(field.Name) && !unsafeBoundary && CanRemainNativeTaggedField(item, field, format)) { WriteTag(builder, field.Name, field.Value, lineEnding, cancellationToken); report.Add("BIBCONV013", BibliographyDiagnosticSeverity.Information, $"Preserved native {format} tag '{field.Name}'.", BibliographyConversionAction.PreservedExtension, item, field.Name); }
        else if (field.Format != format) report.Add("BIBCONV113", BibliographyDiagnosticSeverity.Warning, $"Native {field.Format} field '{field.Name}' cannot be represented safely in {format}.", BibliographyConversionAction.Omitted, item, field.Name);
        else report.Add("BIBCONV122", BibliographyDiagnosticSeverity.Warning, $"Native {format} field '{field.Name}' conflicts with a typed tag or has an unsafe name.", BibliographyConversionAction.Omitted, item, field.Name);
    }

    private static bool CanRemainNativeTaggedField(BibliographyItem item, BibliographyNativeField field, BibliographyFormat format) {
        string tag = field.Name.ToUpperInvariant();
        if (format == BibliographyFormat.Ris) {
            switch (tag) {
                case "ID": return true;
                case "TI": case "T1": return item.Title != null;
                case "T2": case "JF": case "JO": case "JA": return item.ContainerTitle != null;
                case "PB": return item.Publisher != null;
                case "CY": return item.PublisherPlace != null;
                case "ET": return item.Edition != null;
                case "VL": return item.Volume != null;
                case "IS": return item.Issue != null;
                case "AB": case "N2": return item.Abstract != null;
                case "LA": return item.Language != null;
                case "UR": case "L1": return item.Url != null;
                case "SP": case "EP":
                    GetRisPageOutput(item, out bool writeStart, out _, out bool writeEnd, out _);
                    return tag == "SP" ? writeStart : writeEnd;
                case "AU": case "A1": case "ED": case "A2":
                case "PY": case "Y1": case "DA": case "Y2": return false;
                case "DO": case "SN": case "AN": return string.IsNullOrWhiteSpace(field.Value);
                case "KW": case "N1":
                case "TY": case "ER": return false;
                default: return true;
            }
        }
        switch (tag) {
            case "TI": return item.Title != null;
            case "JT": case "TA": return item.ContainerTitle != null;
            case "VI": return item.Volume != null;
            case "IP": return item.Issue != null;
            case "PG": return item.Pages != null;
            case "AB": return item.Abstract != null;
            case "LA": return item.Language != null;
            case "LID": case "AID": return !WouldBindNbibIdentifier(field.Value);
            case "MH": return true;
            case "IS": return string.IsNullOrWhiteSpace(field.Value);
            case "PMID": case "PT": case "FAU": case "AU": case "CN": case "DP": case "OT": case "GN": return false;
            default: return true;
        }
    }

    private static bool WouldBindNbibIdentifier(string value) {
        int marker = value.LastIndexOf(" [", StringComparison.Ordinal);
        if (marker <= 0 || !value.EndsWith("]", StringComparison.Ordinal)) return false;
        string scheme = value.Substring(marker + 2, value.Length - marker - 3).Trim();
        string identifierValue = value.Substring(0, marker).Trim();
        return !string.IsNullOrWhiteSpace(scheme) && !string.IsNullOrWhiteSpace(identifierValue);
    }
    private static bool IsTag(string name) => name.Length >= 2 && name.Length <= 5 && name.All(character => char.IsLetterOrDigit(character));
    private static bool IsRisType(string? value) => !string.IsNullOrWhiteSpace(value) && value!.Length >= 2 && value.Length <= 6 && value.All(char.IsLetterOrDigit);
    private static void AddDocumentNativeLoss(BibliographyDocument document, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        foreach (BibliographyNativeEntry entry in Cancellable(document.NativeEntries, cancellationToken))
            report.Add("BIBCONV114", BibliographyDiagnosticSeverity.Warning, $"Document-level {entry.Format} entry '{entry.Kind}' cannot be represented in {format}.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
    }
    private static IEnumerable<T> Cancellable<T>(IEnumerable<T> source, CancellationToken cancellationToken) {
        int index = 0;
        foreach (T value in source) {
            if ((index++ & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            yield return value;
        }
        cancellationToken.ThrowIfCancellationRequested();
    }
}
