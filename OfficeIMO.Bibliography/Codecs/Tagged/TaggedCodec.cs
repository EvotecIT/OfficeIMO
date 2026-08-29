namespace OfficeIMO.Bibliography;

internal static class TaggedCodec {
    internal static IList<BibliographyItem> ParseRis(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, CancellationToken cancellationToken) =>
        Parse(source, BibliographyFormat.Ris, options, diagnostics, cancellationToken);

    internal static IList<BibliographyItem> ParseNbib(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, CancellationToken cancellationToken) =>
        Parse(source, BibliographyFormat.Nbib, options, diagnostics, cancellationToken);

    internal static string WriteRis(BibliographyDocument document, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        var builder = new StringBuilder();
        for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
            BibliographyItem item = document.Items[itemIndex];
            cancellationToken.ThrowIfCancellationRequested();
            WriteTag(builder, "TY", item.Type == BibliographyItemType.Unknown && IsRisType(item.NativeType) ? item.NativeType!.ToUpperInvariant() : CodecMappings.ToRisType(item.Type), options.LineEnding);
            WriteTag(builder, "ID", CodecMappings.OutputKey(item, itemIndex), options.LineEnding); WriteTag(builder, "TI", item.Title, options.LineEnding); WriteTag(builder, "T2", item.ContainerTitle, options.LineEnding);
            foreach (BibliographyContributor author in item.Contributors.Where(static contributor => contributor.Role == BibliographyContributorRole.Author)) WriteTag(builder, "AU", CodecMappings.FormatName(author.Name), options.LineEnding);
            foreach (BibliographyContributor editor in item.Contributors.Where(static contributor => contributor.Role == BibliographyContributorRole.Editor)) WriteTag(builder, "ED", CodecMappings.FormatName(editor.Name), options.LineEnding);
            WriteDateTags(builder, item, options.LineEnding, "PY", "Y2");
            WriteTag(builder, "PB", item.Publisher, options.LineEnding); WriteTag(builder, "CY", item.PublisherPlace, options.LineEnding); WriteTag(builder, "ET", item.Edition, options.LineEnding);
            WriteTag(builder, "VL", item.Volume, options.LineEnding); WriteTag(builder, "IS", item.Issue, options.LineEnding); WritePages(builder, item.Pages, options.LineEnding);
            WriteTag(builder, "AB", item.Abstract, options.LineEnding); WriteTag(builder, "LA", item.Language, options.LineEnding); WriteTag(builder, "UR", item.Url, options.LineEnding);
            foreach (BibliographyIdentifier identifier in item.Identifiers) WriteRisIdentifier(builder, identifier, options.LineEnding);
            foreach (string keyword in item.Keywords) WriteTag(builder, "KW", keyword, options.LineEnding);
            foreach (string note in item.Notes) WriteTag(builder, "N1", note, options.LineEnding);
            WriteNativeFields(builder, item, BibliographyFormat.Ris, options.LineEnding, report);
            WriteTag(builder, "ER", string.Empty, options.LineEnding);
            if (itemIndex + 1 < document.Items.Count) builder.Append(options.LineEnding);
        }
        AddDocumentNativeLoss(document, BibliographyFormat.Ris, report);
        return builder.ToString();
    }

    internal static string WriteNbib(BibliographyDocument document, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        var builder = new StringBuilder();
        for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
            BibliographyItem item = document.Items[itemIndex];
            cancellationToken.ThrowIfCancellationRequested();
            WriteTag(builder, "PMID", item.GetIdentifier("PMID") ?? CodecMappings.OutputKey(item, itemIndex), options.LineEnding);
            WriteTag(builder, "TI", item.Title, options.LineEnding); WriteTag(builder, "JT", item.ContainerTitle, options.LineEnding);
            foreach (BibliographyContributor author in item.Contributors.Where(static contributor => contributor.Role == BibliographyContributorRole.Author && string.IsNullOrWhiteSpace(contributor.Name.Literal))) WriteTag(builder, "FAU", CodecMappings.FormatName(author.Name), options.LineEnding);
            foreach (BibliographyContributor author in item.Contributors.Where(static contributor => contributor.Role == BibliographyContributorRole.Author && string.IsNullOrWhiteSpace(contributor.Name.Literal))) WriteTag(builder, "AU", CompactName(author.Name), options.LineEnding);
            foreach (BibliographyContributor author in item.Contributors.Where(static contributor => contributor.Role == BibliographyContributorRole.Author && !string.IsNullOrWhiteSpace(contributor.Name.Literal))) WriteTag(builder, "CN", author.Name.Literal, options.LineEnding);
            BibliographyDate? issued = item.GetDate(BibliographyDateRole.Issued); if (issued != null) WriteTag(builder, "DP", CodecMappings.FormatDate(issued), options.LineEnding);
            WriteTag(builder, "VI", item.Volume, options.LineEnding); WriteTag(builder, "IP", item.Issue, options.LineEnding); WriteTag(builder, "PG", item.Pages, options.LineEnding);
            WriteTag(builder, "AB", item.Abstract, options.LineEnding); WriteTag(builder, "LA", item.Language, options.LineEnding);
            foreach (BibliographyIdentifier identifier in item.Identifiers) WriteNbibIdentifier(builder, identifier, options.LineEnding);
            foreach (string keyword in item.Keywords) WriteTag(builder, "OT", keyword, options.LineEnding);
            foreach (string note in item.Notes) WriteTag(builder, "GN", note, options.LineEnding);
            WriteNativeFields(builder, item, BibliographyFormat.Nbib, options.LineEnding, report);
            if (itemIndex + 1 < document.Items.Count) builder.Append(options.LineEnding);
        }
        AddDocumentNativeLoss(document, BibliographyFormat.Nbib, report);
        return builder.ToString();
    }

    private static IList<BibliographyItem> Parse(string source, BibliographyFormat format, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, CancellationToken cancellationToken) {
        var items = new List<BibliographyItem>();
        var limits = new BibliographyLimitGuard(options);
        var diagnosticGuard = new BibliographyDiagnosticGuard(options, diagnostics, items);
        BibliographyItem? current = null;
        string? previousTag = null;
        BibliographyNativeField? previousNativeField = null;
        string[] lines = source.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            string line = lines[lineIndex];
            if (string.IsNullOrWhiteSpace(line)) {
                if (format == BibliographyFormat.Nbib) current = null;
                previousTag = null; previousNativeField = null; continue;
            }
            if (TrySplitLine(line, format, out string tag, out string value)) {
                if (format == BibliographyFormat.Ris && string.Equals(tag, "TY", StringComparison.OrdinalIgnoreCase)) {
                    limits.AddValue(items, value, lineIndex);
                    current = NewItem(items, limits, lineIndex); current.NativeType = value; current.Type = CodecMappings.ParseType(value);
                    previousTag = tag;
                    previousNativeField = null;
                    continue;
                } else if (format == BibliographyFormat.Nbib && string.Equals(tag, "PMID", StringComparison.OrdinalIgnoreCase) && current != null && !string.IsNullOrEmpty(current.Key)) {
                    current = NewItem(items, limits, lineIndex);
                } else if (current == null) current = NewItem(items, limits, lineIndex);
                if (format == BibliographyFormat.Nbib && current!.Type == BibliographyItemType.Unknown) { current.Type = BibliographyItemType.ArticleJournal; current.NativeType = "Journal Article"; }
                limits.AddValue(items, value, lineIndex);
                if (format == BibliographyFormat.Ris && string.Equals(tag, "ER", StringComparison.OrdinalIgnoreCase)) current = null;
                else {
                    int nativeCount = current!.NativeFields.Count;
                    Bind(current, format, tag, value);
                    previousNativeField = current.NativeFields.Count > nativeCount ? current.NativeFields[current.NativeFields.Count - 1] : null;
                }
                previousTag = tag;
            } else if (current != null && previousTag != null && (char.IsWhiteSpace(line[0]) || format == BibliographyFormat.Ris)) {
                string continuation = line.Trim();
                limits.AddValue(items, continuation, lineIndex);
                if (previousNativeField != null) previousNativeField.Value = AppendChecked(previousNativeField.Value, continuation, items, limits, lineIndex + 1);
                else AppendContinuation(current, format, previousTag, continuation, diagnosticGuard, lineIndex + 1, items, limits);
            } else diagnosticGuard.Add(new BibliographyDiagnostic("BIBTAG001", BibliographyDiagnosticSeverity.Warning, $"Ignored malformed {format} line.", line: lineIndex + 1, column: 1));
        }
        if (format == BibliographyFormat.Nbib) NormalizeNbibAuthors(items);
        foreach (BibliographyItem item in items.Where(static item => string.IsNullOrWhiteSpace(item.Key))) diagnosticGuard.Add(new BibliographyDiagnostic("BIBTAG003", BibliographyDiagnosticSeverity.Warning, $"{format} record has no citation identifier.", itemKey: item.Key));
        return items;
    }

    private static BibliographyItem NewItem(IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) { limits.AddItem(items, offset); var item = new BibliographyItem(); items.Add(item); return item; }

    private static bool TrySplitLine(string line, BibliographyFormat format, out string tag, out string value) {
        tag = string.Empty; value = string.Empty;
        if (line.Length == 0 || char.IsWhiteSpace(line[0])) return false;
        int dashPosition = line.IndexOf('-');
        if (dashPosition <= 0) return false;
        tag = line.Substring(0, dashPosition).Trim();
        if (tag.Length < 2 || tag.Length > 5 || !tag.All(char.IsLetterOrDigit)) { tag = string.Empty; return false; }
        value = line.Substring(dashPosition + 1).TrimStart();
        return true;
    }

    private static void Bind(BibliographyItem item, BibliographyFormat format, string tag, string value) {
        string field = tag.ToUpperInvariant();
        if (format == BibliographyFormat.Ris) BindRis(item, field, value); else BindNbib(item, field, value);
    }

    private static void BindRis(BibliographyItem item, string field, string value) {
        switch (field) {
            case "ID": SetScalar(item, BibliographyFormat.Ris, "key", field, value, assigned => item.Key = assigned); break;
            case "TI": case "T1": SetScalar(item, BibliographyFormat.Ris, "title", field, value, assigned => item.Title = assigned); break;
            case "T2": case "JF": case "JO": case "JA": SetScalar(item, BibliographyFormat.Ris, "container-title", field, value, assigned => item.ContainerTitle = assigned); break;
            case "AU": case "A1": item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, CodecMappings.ParseCommaName(value))); break;
            case "ED": case "A2": item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Editor, CodecMappings.ParseCommaName(value))); break;
            case "PY": case "Y1": case "DA": item.Dates.Add(CodecMappings.ParseDate(BibliographyDateRole.Issued, value)); break;
            case "Y2": item.Dates.Add(CodecMappings.ParseDate(BibliographyDateRole.Accessed, value)); break;
            case "PB": SetScalar(item, BibliographyFormat.Ris, "publisher", field, value, assigned => item.Publisher = assigned); break;
            case "CY": SetScalar(item, BibliographyFormat.Ris, "publisher-place", field, value, assigned => item.PublisherPlace = assigned); break;
            case "ET": SetScalar(item, BibliographyFormat.Ris, "edition", field, value, assigned => item.Edition = assigned); break;
            case "VL": SetScalar(item, BibliographyFormat.Ris, "volume", field, value, assigned => item.Volume = assigned); break;
            case "IS": SetScalar(item, BibliographyFormat.Ris, "issue", field, value, assigned => item.Issue = assigned); break;
            case "SP": SetScalar(item, BibliographyFormat.Ris, "pages-start", field, value, assigned => item.Pages = assigned); break;
            case "EP": SetPageEnd(item, value); break;
            case "AB": case "N2": SetScalar(item, BibliographyFormat.Ris, "abstract", field, value, assigned => item.Abstract = assigned); break;
            case "LA": SetScalar(item, BibliographyFormat.Ris, "language", field, value, assigned => item.Language = assigned); break;
            case "UR": case "L1": SetScalar(item, BibliographyFormat.Ris, "url", field, value, assigned => item.Url = assigned); break;
            case "DO": CodecMappings.AddIdentifier(item, "DOI", value); break; case "SN": CodecMappings.AddIdentifier(item, CodecMappings.InferSerialScheme(value), value); break;
            case "AN": if (string.IsNullOrWhiteSpace(item.Key)) item.Key = value; ParseRisAccession(item, value); break;
            case "KW": item.Keywords.Add(value); break; case "N1": item.Notes.Add(value); break;
            default: item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Ris, field, value)); break;
        }
    }

    private static void BindNbib(BibliographyItem item, string field, string value) {
        switch (field) {
            case "PMID": item.Key = value; CodecMappings.AddIdentifier(item, "PMID", value); break;
            case "TI": SetScalar(item, BibliographyFormat.Nbib, "title", field, value, assigned => item.Title = assigned); break;
            case "JT": case "TA": SetScalar(item, BibliographyFormat.Nbib, "container-title", field, value, assigned => item.ContainerTitle = assigned); break;
            case "FAU": item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, CodecMappings.ParseCommaName(value))); break;
            case "AU": item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Nbib, field, value)); break;
            case "CN": item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = value })); break;
            case "DP": item.Dates.Add(CodecMappings.ParseDate(BibliographyDateRole.Issued, value)); break;
            case "VI": SetScalar(item, BibliographyFormat.Nbib, "volume", field, value, assigned => item.Volume = assigned); break;
            case "IP": SetScalar(item, BibliographyFormat.Nbib, "issue", field, value, assigned => item.Issue = assigned); break;
            case "PG": SetScalar(item, BibliographyFormat.Nbib, "pages", field, value, assigned => item.Pages = assigned); break;
            case "AB": SetScalar(item, BibliographyFormat.Nbib, "abstract", field, value, assigned => item.Abstract = assigned); break;
            case "LA": SetScalar(item, BibliographyFormat.Nbib, "language", field, value, assigned => item.Language = assigned); break;
            case "LID": case "AID": ParseNbibIdentifier(item, value, field); break;
            case "IS": CodecMappings.AddIdentifier(item, "ISSN", StripBracketQualifier(value)); break;
            case "OT": case "MH": item.Keywords.Add(value); break; case "GN": item.Notes.Add(value); break;
            default: item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Nbib, field, value)); break;
        }
    }

    private static void ParseNbibIdentifier(BibliographyItem item, string value, string field) {
        int marker = value.LastIndexOf(" [", StringComparison.Ordinal);
        if (marker > 0 && value.EndsWith("]", StringComparison.Ordinal)) { string scheme = value.Substring(marker + 2, value.Length - marker - 3); CodecMappings.AddIdentifier(item, scheme, value.Substring(0, marker)); }
        else item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Nbib, field, value));
    }

    private static void ParseRisAccession(BibliographyItem item, string value) {
        int separator = value.IndexOf(':');
        if (separator > 0 && separator + 1 < value.Length) CodecMappings.AddIdentifier(item, value.Substring(0, separator), value.Substring(separator + 1));
        else CodecMappings.AddIdentifier(item, "accession", value);
    }

    private static void SetScalar(BibliographyItem item, BibliographyFormat format, string semanticName, string sourceTag, string value, Action<string> write) {
        string binding = format + ":" + semanticName;
        if (item.TaggedScalarBindings.Add(binding)) write(value);
        else item.NativeFields.Add(new BibliographyNativeField(format, sourceTag, value));
    }

    private static void SetPageEnd(BibliographyItem item, string value) {
        const string binding = "Ris:pages-end";
        if (item.TaggedScalarBindings.Add(binding)) item.Pages = string.IsNullOrWhiteSpace(item.Pages) ? value : item.Pages + "-" + value;
        else item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Ris, "EP", value));
    }

    private static string StripBracketQualifier(string value) { int marker = value.LastIndexOf(" (", StringComparison.Ordinal); return marker > 0 && value.EndsWith(")", StringComparison.Ordinal) ? value.Substring(0, marker) : value; }

    private static void AppendContinuation(BibliographyItem item, BibliographyFormat format, string tag, string value, BibliographyDiagnosticGuard diagnostics, int line, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        string field = tag.ToUpperInvariant();
        if (format == BibliographyFormat.Ris) {
            switch (field) {
                case "TI": case "T1": item.Title = AppendChecked(item.Title, value, items, limits, line); return;
                case "T2": case "JF": case "JO": case "JA": item.ContainerTitle = AppendChecked(item.ContainerTitle, value, items, limits, line); return;
                case "AB": case "N2": item.Abstract = AppendChecked(item.Abstract, value, items, limits, line); return;
                case "UR": case "L1": item.Url = AppendChecked(item.Url, value, items, limits, line); return;
                case "N1": AppendLast(item.Notes, value, items, limits, line); return;
                case "KW": AppendLast(item.Keywords, value, items, limits, line); return;
                case "AU": case "A1": AppendContributor(item, BibliographyContributorRole.Author, value, items, limits, line); return;
                case "ED": case "A2": AppendContributor(item, BibliographyContributorRole.Editor, value, items, limits, line); return;
            }
        } else {
            switch (field) {
                case "TI": item.Title = AppendChecked(item.Title, value, items, limits, line); return;
                case "JT": case "TA": item.ContainerTitle = AppendChecked(item.ContainerTitle, value, items, limits, line); return;
                case "AB": item.Abstract = AppendChecked(item.Abstract, value, items, limits, line); return;
                case "GN": AppendLast(item.Notes, value, items, limits, line); return;
                case "OT": case "MH": AppendLast(item.Keywords, value, items, limits, line); return;
                case "FAU": AppendContributor(item, BibliographyContributorRole.Author, value, items, limits, line); return;
                case "CN": AppendContributor(item, BibliographyContributorRole.Author, value, items, limits, line); return;
            }
        }
        if (item.NativeFields.LastOrDefault(nativeField => nativeField.Format == format && string.Equals(nativeField.Name, tag, StringComparison.OrdinalIgnoreCase)) is BibliographyNativeField native) {
            native.Value = AppendChecked(native.Value, value, items, limits, line);
        } else {
            item.NativeFields.Add(new BibliographyNativeField(format, tag, value));
            diagnostics.Add(new BibliographyDiagnostic("BIBTAG002", BibliographyDiagnosticSeverity.Information, $"Continuation for '{tag}' was retained as a native field.", line: line, column: 1, itemKey: item.Key, field: tag));
        }
    }

    private static string Append(string? current, string value) => string.IsNullOrEmpty(current) ? value : current + " " + value;
    private static string AppendChecked(string? current, string value, IList<BibliographyItem> items, BibliographyLimitGuard limits, int offset) { string combined = Append(current, value); limits.CheckValueLength(items, combined, offset); return combined; }
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

    private static void NormalizeNbibAuthors(IEnumerable<BibliographyItem> items) {
        foreach (BibliographyItem item in items) {
            BibliographyNativeField[] compactAuthors = item.NativeFields.Where(field => field.Format == BibliographyFormat.Nbib && string.Equals(field.Name, "AU", StringComparison.OrdinalIgnoreCase)).ToArray();
            int fullAuthorCount = item.Contributors.Count(static contributor => contributor.Role == BibliographyContributorRole.Author && string.IsNullOrWhiteSpace(contributor.Name.Literal));
            int firstAdditional = fullAuthorCount == 0 ? 0 : Math.Min(fullAuthorCount, compactAuthors.Length);
            for (int index = firstAdditional; index < compactAuthors.Length; index++) item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, CodecMappings.ParseCommaName(compactAuthors[index].Value)));
            foreach (BibliographyNativeField field in compactAuthors) item.NativeFields.Remove(field);
        }
    }
    private static void WriteTag(StringBuilder builder, string tag, string? value, string lineEnding) { if (value == null) return; string prefix = tag.PadRight(4) + "- "; string[] lines = value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'); builder.Append(prefix).Append(lines[0]).Append(lineEnding); for (int index = 1; index < lines.Length; index++) builder.Append("      ").Append(lines[index]).Append(lineEnding); }
    private static void WritePages(StringBuilder builder, string? pages, string lineEnding) { if (string.IsNullOrWhiteSpace(pages)) return; string[] parts = pages!.Split(new[] { '-' }, 2); WriteTag(builder, "SP", parts[0], lineEnding); if (parts.Length > 1) WriteTag(builder, "EP", parts[1], lineEnding); }
    private static void WriteDateTags(StringBuilder builder, BibliographyItem item, string lineEnding, string issuedTag, string accessedTag) { BibliographyDate? issued = item.GetDate(BibliographyDateRole.Issued); if (issued != null) WriteTag(builder, issuedTag, CodecMappings.FormatDate(issued), lineEnding); BibliographyDate? accessed = item.GetDate(BibliographyDateRole.Accessed); if (accessed != null) WriteTag(builder, accessedTag, CodecMappings.FormatDate(accessed), lineEnding); }
    private static void WriteRisIdentifier(StringBuilder builder, BibliographyIdentifier identifier, string lineEnding) { if (string.Equals(identifier.Scheme, "DOI", StringComparison.OrdinalIgnoreCase)) WriteTag(builder, "DO", identifier.Value, lineEnding); else if (string.Equals(identifier.Scheme, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "SN", StringComparison.OrdinalIgnoreCase)) WriteTag(builder, "SN", identifier.Value, lineEnding); else if (string.Equals(identifier.Scheme, "accession", StringComparison.OrdinalIgnoreCase)) WriteTag(builder, "AN", identifier.Value.IndexOf(':') >= 0 ? "accession:" + identifier.Value : identifier.Value, lineEnding); else WriteTag(builder, "AN", identifier.Scheme + ":" + identifier.Value, lineEnding); }
    private static void WriteNbibIdentifier(StringBuilder builder, BibliographyIdentifier identifier, string lineEnding) { if (string.Equals(identifier.Scheme, "PMID", StringComparison.OrdinalIgnoreCase)) return; if (string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase)) WriteTag(builder, "IS", identifier.Value, lineEnding); else WriteTag(builder, "AID", identifier.Value + " [" + identifier.Scheme.ToLowerInvariant() + "]", lineEnding); }
    private static string CompactName(BibliographyName name) => string.IsNullOrWhiteSpace(name.Literal) ? ((name.Family ?? string.Empty) + " " + Initials(name.Given)).Trim() : name.Literal!;
    private static string Initials(string? value) => string.IsNullOrWhiteSpace(value) ? string.Empty : string.Concat(value!.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries).Select(static part => part.Substring(0, 1)));

    private static void WriteNativeFields(StringBuilder builder, BibliographyItem item, BibliographyFormat format, string lineEnding, BibliographyConversionReport report) {
        foreach (BibliographyNativeField field in item.NativeFields) {
            bool unsafeBoundary = format == BibliographyFormat.Ris && (string.Equals(field.Name, "TY", StringComparison.OrdinalIgnoreCase) || string.Equals(field.Name, "ER", StringComparison.OrdinalIgnoreCase)) || format == BibliographyFormat.Nbib && string.Equals(field.Name, "PMID", StringComparison.OrdinalIgnoreCase);
            if (field.Format == format && IsTag(field.Name) && !unsafeBoundary) { WriteTag(builder, field.Name.ToUpperInvariant(), field.Value, lineEnding); report.Add("BIBCONV013", BibliographyDiagnosticSeverity.Information, $"Preserved native {format} tag '{field.Name}'.", BibliographyConversionAction.PreservedExtension, item, field.Name); }
            else if (field.Format != format) report.Add("BIBCONV113", BibliographyDiagnosticSeverity.Warning, $"Native {field.Format} field '{field.Name}' cannot be represented safely in {format}.", BibliographyConversionAction.Omitted, item, field.Name);
            else report.Add("BIBCONV122", BibliographyDiagnosticSeverity.Warning, $"Native {format} field '{field.Name}' conflicts with a typed tag or has an unsafe name.", BibliographyConversionAction.Omitted, item, field.Name);
        }
    }
    private static bool IsTag(string name) => name.Length >= 2 && name.Length <= 5 && name.All(character => char.IsLetterOrDigit(character));
    private static bool IsRisType(string? value) => !string.IsNullOrWhiteSpace(value) && value!.Length >= 2 && value.Length <= 6 && value.All(char.IsLetterOrDigit);
    private static void AddDocumentNativeLoss(BibliographyDocument document, BibliographyFormat format, BibliographyConversionReport report) { foreach (BibliographyNativeEntry entry in document.NativeEntries) report.Add("BIBCONV114", BibliographyDiagnosticSeverity.Warning, $"Document-level {entry.Format} entry '{entry.Kind}' cannot be represented in {format}.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind); }
}
