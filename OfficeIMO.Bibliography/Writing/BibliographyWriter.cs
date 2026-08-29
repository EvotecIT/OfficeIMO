namespace OfficeIMO.Bibliography;

internal static class BibliographyWriter {
    internal static BibliographyWriteResult Write(BibliographyDocument document, BibliographyWriteOptions? options, CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new BibliographyWriteOptions();
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        BibliographyFormat format = options.Format ?? document.SourceFormat;
        var report = new BibliographyConversionReport();

        if (options.Mode == BibliographyWriterMode.Preserve && format == document.SourceFormat && !document.IsModified && document.OriginalText != null) {
            if (document.OriginalBytes == null) InspectEncoding(document.OriginalText, options.Encoding, report);
            if (options.RequireNoLoss) report.RequireNoLoss();
            byte[] bytes = document.OriginalBytes != null ? (byte[])document.OriginalBytes.Clone() : Encode(document.OriginalText, options.Encoding);
            return new BibliographyWriteResult(document.OriginalText, bytes, format, true, report);
        }

        BibliographyConversionInspector.Inspect(document, format, report);
        string content;
        switch (format) {
            case BibliographyFormat.BibTex:
            case BibliographyFormat.BibLatex:
                content = BibCodec.Write(document, format, options, report, cancellationToken);
                break;
            case BibliographyFormat.CslJson:
                content = CslJsonCodec.Write(document, options, report, cancellationToken);
                break;
            case BibliographyFormat.Ris:
                content = TaggedCodec.WriteRis(document, options, report, cancellationToken);
                break;
            case BibliographyFormat.Nbib:
                content = TaggedCodec.WriteNbib(document, options, report, cancellationToken);
                break;
            case BibliographyFormat.EndNoteXml:
                content = EndNoteXmlCodec.Write(document, options, report, cancellationToken);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(options), format, "Unknown bibliography format.");
        }
        InspectEncoding(content, options.Encoding, report);
        if (options.RequireNoLoss) report.RequireNoLoss();
        return new BibliographyWriteResult(content, Encode(content, options.Encoding), format, false, report);
    }

    private static byte[] Encode(string value, Encoding encoding) {
        byte[] preamble = encoding.GetPreamble();
        byte[] content = encoding.GetBytes(value);
        if (preamble.Length == 0) return content;
        var result = new byte[preamble.Length + content.Length];
        Buffer.BlockCopy(preamble, 0, result, 0, preamble.Length); Buffer.BlockCopy(content, 0, result, preamble.Length, content.Length);
        return result;
    }

    private static void InspectEncoding(string value, Encoding encoding, BibliographyConversionReport report) {
        var strictEncoding = (Encoding)encoding.Clone();
        strictEncoding.EncoderFallback = EncoderFallback.ExceptionFallback;
        try {
            strictEncoding.GetByteCount(value);
        } catch (EncoderFallbackException) {
            report.Add("BIBCONV220", BibliographyDiagnosticSeverity.Warning, $"The selected {encoding.WebName} encoding cannot represent all output characters without replacement.", BibliographyConversionAction.Approximated, field: "encoding");
        }
    }
}

internal static class BibliographyConversionInspector {
    internal static void Inspect(BibliographyDocument document, BibliographyFormat format, BibliographyConversionReport report) {
        foreach (BibliographyDiagnostic diagnostic in document.Diagnostics.Where(IsRecoveryLossDiagnostic)) {
            report.Add("BIBCONV222", BibliographyDiagnosticSeverity.Warning, $"Canonical output is based on partially recovered source after parser diagnostic {diagnostic.Code}; unrecovered source content may be omitted.", BibliographyConversionAction.Omitted, field: diagnostic.Field);
        }
        InspectKeys(document, format, report);
        foreach (BibliographyItem item in document.Items) {
            InspectType(item, document.SourceFormat, format, report); InspectContributors(item, format, report); InspectDates(item, format, report); InspectNestedNativeFields(item, format, report); InspectProperties(item, format, report); InspectIdentifiers(item, format, report); InspectRepeatableValues(item, format, report); InspectTextEncoding(item, format, report);
        }
    }

    private static bool IsRecoveryLossDiagnostic(BibliographyDiagnostic diagnostic) =>
        diagnostic.Severity == BibliographyDiagnosticSeverity.Error || diagnostic.Code == "BIBBIB001" || diagnostic.Code == "BIBTAG001" || diagnostic.Code == "BIBCSL003";

    private static void InspectKeys(BibliographyDocument document, BibliographyFormat format, BibliographyConversionReport report) {
        foreach (BibliographyItem item in document.Items.Where(static item => string.IsNullOrWhiteSpace(item.Key)))
            Loss(report, item, "key", "BIBCONV215", $"A missing citation key is replaced with a deterministic generated identifier in {format}.", BibliographyConversionAction.Approximated);
        foreach (IGrouping<string, BibliographyItem> duplicate in document.Items.Where(static item => !string.IsNullOrWhiteSpace(item.Key)).GroupBy(static item => item.Key, StringComparer.OrdinalIgnoreCase).Where(static group => group.Count() > 1)) {
            foreach (BibliographyItem item in duplicate) Loss(report, item, "key", "BIBCONV216", $"Duplicate citation key '{duplicate.Key}' is not unique in {format} output.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex) {
            foreach (BibliographyItem item in document.Items.Where(static item => !string.IsNullOrWhiteSpace(item.Key) && item.Key.Any(character => !BibCodec.IsSafeKeyCharacter(character))))
                Loss(report, item, "key", "BIBCONV217", "The citation key contains characters that must be normalized for BibTeX output.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.Nbib) {
            foreach (BibliographyItem item in document.Items) {
                BibliographyIdentifier[] pmids = item.Identifiers.Where(static identifier => string.Equals(identifier.Scheme, "PMID", StringComparison.OrdinalIgnoreCase)).ToArray();
                string? pmid = pmids.FirstOrDefault()?.Value;
                if (string.IsNullOrWhiteSpace(pmid)) Loss(report, item, "identifiers.PMID", "BIBCONV223", "NBIB output uses the citation key as a PMID because the item has no PMID identifier.", BibliographyConversionAction.Approximated);
                else if (!string.Equals(item.Key, pmid, StringComparison.Ordinal)) Loss(report, item, "key", "BIBCONV224", "NBIB represents PMID as its record identifier, so the distinct citation key is omitted.", BibliographyConversionAction.Omitted);
                if (pmids.Length > 1) Loss(report, item, "identifiers.PMID", "BIBCONV227", "NBIB represents one PMID as record identity, so additional PMID identifiers are omitted.", BibliographyConversionAction.Omitted);
            }
        }
    }

    private static void InspectType(BibliographyItem item, BibliographyFormat sourceFormat, BibliographyFormat format, BibliographyConversionReport report) {
        bool exact;
        bool sameFormatNativeType = item.Type == BibliographyItemType.Unknown && !string.IsNullOrWhiteSpace(item.NativeType) && sourceFormat == format;
        switch (format) {
            case BibliographyFormat.CslJson:
                exact = sameFormatNativeType || item.Type != BibliographyItemType.Unknown && item.Type != BibliographyItemType.Proceedings && item.Type != BibliographyItemType.Document;
                break;
            case BibliographyFormat.BibTex: case BibliographyFormat.BibLatex:
                bool hasNativeBibType = (sourceFormat == BibliographyFormat.BibTex || sourceFormat == BibliographyFormat.BibLatex) && !string.IsNullOrWhiteSpace(item.NativeType);
                exact = hasNativeBibType ? BibCodec.CanPreserveNativeType(sourceFormat, format, item) : item.Type == BibliographyItemType.ArticleJournal || item.Type == BibliographyItemType.Book || item.Type == BibliographyItemType.Chapter || item.Type == BibliographyItemType.PaperConference || item.Type == BibliographyItemType.Proceedings || item.Type == BibliographyItemType.Report || item.Type == BibliographyItemType.Thesis || item.Type == BibliographyItemType.Manuscript;
                break;
            case BibliographyFormat.Ris:
                exact = sameFormatNativeType && IsSafeRisType(item.NativeType) || item.Type != BibliographyItemType.Unknown && item.Type != BibliographyItemType.Article && item.Type != BibliographyItemType.Proceedings && item.Type != BibliographyItemType.LegalCase && item.Type != BibliographyItemType.Manuscript && item.Type != BibliographyItemType.Document;
                break;
            case BibliographyFormat.Nbib:
                exact = item.Type == BibliographyItemType.ArticleJournal || sourceFormat == BibliographyFormat.Nbib && item.NativeFields.Any(field => field.Format == BibliographyFormat.Nbib && string.Equals(field.Name, "PT", StringComparison.OrdinalIgnoreCase) && CodecMappings.ParseType(field.Value) == item.Type);
                break;
            case BibliographyFormat.EndNoteXml:
                exact = item.Type == BibliographyItemType.ArticleJournal || item.Type == BibliographyItemType.Book || item.Type == BibliographyItemType.Chapter || item.Type == BibliographyItemType.PaperConference || item.Type == BibliographyItemType.Report || item.Type == BibliographyItemType.Thesis || item.Type == BibliographyItemType.WebPage || item.Type == BibliographyItemType.Patent;
                break;
            default: exact = false; break;
        }
        if (!exact) Loss(report, item, "type", "BIBCONV200", $"Item type '{item.Type}' is written using a broader {format} type.", BibliographyConversionAction.Approximated);
    }

    private static bool IsSafeRisType(string? value) => !string.IsNullOrWhiteSpace(value) && value!.Length >= 2 && value.Length <= 6 && value.All(char.IsLetterOrDigit);

    private static void InspectContributors(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report) {
        foreach (BibliographyContributorRole role in item.Contributors.Select(static value => value.Role).Distinct()) {
            bool exact;
            switch (format) {
                case BibliographyFormat.BibTex: exact = role == BibliographyContributorRole.Author || role == BibliographyContributorRole.Editor; break;
                case BibliographyFormat.BibLatex: exact = role == BibliographyContributorRole.Author || role == BibliographyContributorRole.Editor || role == BibliographyContributorRole.Translator; break;
                case BibliographyFormat.CslJson: exact = role != BibliographyContributorRole.Other; break;
                case BibliographyFormat.Ris: exact = role == BibliographyContributorRole.Author || role == BibliographyContributorRole.Editor; break;
                case BibliographyFormat.Nbib: exact = role == BibliographyContributorRole.Author; break;
                case BibliographyFormat.EndNoteXml: exact = role == BibliographyContributorRole.Author || role == BibliographyContributorRole.Editor || role == BibliographyContributorRole.CollectionEditor || role == BibliographyContributorRole.Translator; break;
                default: exact = false; break;
            }
            if (!exact) Loss(report, item, "contributors." + role, "BIBCONV201", $"Contributor role '{role}' is not represented exactly in {format}.", format == BibliographyFormat.EndNoteXml ? BibliographyConversionAction.Approximated : BibliographyConversionAction.Omitted);
        }
        if (format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex) {
            foreach (BibliographyContributor contributor in item.Contributors.Where(static contributor => !BibCodec.CanRoundTripStructuredName(contributor.Name)))
                Loss(report, item, "contributors", "BIBCONV226", "A structured contributor particle does not follow BibTeX lowercase-particle syntax and cannot be reopened exactly.", BibliographyConversionAction.Approximated);
        } else if (format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib || format == BibliographyFormat.EndNoteXml) {
            foreach (BibliographyContributor contributor in item.Contributors.Where(static contributor => !string.IsNullOrWhiteSpace(contributor.Name.DroppingParticle) || !string.IsNullOrWhiteSpace(contributor.Name.NonDroppingParticle)))
                Loss(report, item, "contributors", "BIBCONV229", $"Structured contributor particles are flattened in {format} output and cannot be reopened exactly.", BibliographyConversionAction.Approximated);
        }
        if (ReordersContributors(item, format))
            Loss(report, item, "contributors", "BIBCONV230", $"Contributor source order is regrouped by {format} output and cannot be reopened exactly.", BibliographyConversionAction.Approximated);
    }

    private static bool ReordersContributors(BibliographyItem item, BibliographyFormat format) {
        BibliographyContributor[] source;
        BibliographyContributor[] output;
        switch (format) {
            case BibliographyFormat.BibTex: case BibliographyFormat.BibLatex:
                BibliographyContributorRole[] bibRoles = { BibliographyContributorRole.Author, BibliographyContributorRole.Editor, BibliographyContributorRole.Translator };
                source = item.Contributors.Where(contributor => bibRoles.Contains(contributor.Role)).ToArray();
                output = bibRoles.SelectMany(role => source.Where(contributor => contributor.Role == role)).ToArray();
                break;
            case BibliographyFormat.CslJson:
                BibliographyContributorRole[] cslRoles = { BibliographyContributorRole.Author, BibliographyContributorRole.Editor, BibliographyContributorRole.Translator, BibliographyContributorRole.Recipient, BibliographyContributorRole.Interviewer, BibliographyContributorRole.Composer, BibliographyContributorRole.CollectionEditor };
                source = item.Contributors.Where(contributor => cslRoles.Contains(contributor.Role)).ToArray();
                output = cslRoles.SelectMany(role => source.Where(contributor => contributor.Role == role)).ToArray();
                break;
            case BibliographyFormat.EndNoteXml:
                BibliographyContributorRole[] endNoteRoles = { BibliographyContributorRole.Author, BibliographyContributorRole.Editor, BibliographyContributorRole.CollectionEditor, BibliographyContributorRole.Translator };
                source = item.Contributors.Where(contributor => endNoteRoles.Contains(contributor.Role)).ToArray();
                output = source.GroupBy(static contributor => contributor.Role).SelectMany(static group => group).ToArray();
                break;
            case BibliographyFormat.Nbib:
                source = item.Contributors.Where(static contributor => contributor.Role == BibliographyContributorRole.Author).ToArray();
                output = source.Where(static contributor => string.IsNullOrWhiteSpace(contributor.Name.Literal)).Concat(source.Where(static contributor => !string.IsNullOrWhiteSpace(contributor.Name.Literal))).ToArray();
                break;
            default:
                return false;
        }
        return !source.SequenceEqual(output);
    }

    private static void InspectDates(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report) {
        foreach (BibliographyDateRole role in item.Dates.Select(static value => value.Role).Distinct()) {
            bool exact = format == BibliographyFormat.CslJson ? role != BibliographyDateRole.Other
                : (format == BibliographyFormat.BibLatex || format == BibliographyFormat.Ris) ? role == BibliographyDateRole.Issued || role == BibliographyDateRole.Accessed
                : format == BibliographyFormat.BibTex ? role == BibliographyDateRole.Issued
                : role == BibliographyDateRole.Issued;
            if (!exact) Loss(report, item, "dates." + role, "BIBCONV202", $"Date role '{role}' is not represented in {format}.", BibliographyConversionAction.Omitted);
            if (item.Dates.Count(date => date.Role == role) > 1) Loss(report, item, "dates." + role, "BIBCONV205", $"Multiple '{role}' dates collapse to the first value in {format}.", BibliographyConversionAction.Approximated);
        }
        BibliographyDate? issued = item.GetDate(BibliographyDateRole.Issued);
        if (format == BibliographyFormat.BibTex && issued?.Day != null) Loss(report, item, "dates.Issued.day", "BIBCONV212", "Classic BibTeX output omits issued-day precision.", BibliographyConversionAction.Omitted);
        foreach (BibliographyDate date in item.Dates) {
            if (!IsValidDate(date.Year, date.Month, date.Day) || !IsValidDate(date.EndYear, date.EndMonth, date.EndDay) || date.EndYear.HasValue && !date.Year.HasValue)
                Loss(report, item, "dates." + date.Role, "BIBCONV218", "A date contains an invalid or incomplete numeric component sequence.", BibliographyConversionAction.Approximated);
            if (date.EndYear.HasValue && format != BibliographyFormat.CslJson && format != BibliographyFormat.BibLatex && format != BibliographyFormat.EndNoteXml)
                Loss(report, item, "dates." + date.Role + ".end", "BIBCONV219", $"Date ranges are not represented exactly in {format}.", BibliographyConversionAction.Approximated);
            if (format != BibliographyFormat.CslJson && date.Year.HasValue && !string.IsNullOrWhiteSpace(date.Literal))
                Loss(report, item, "dates." + date.Role + ".literal", "BIBCONV221", $"The literal date value is not represented alongside numeric date parts in {format}.", BibliographyConversionAction.Omitted);
        }
    }

    private static bool IsValidDate(int? year, int? month, int? day) {
        if (!year.HasValue) return !month.HasValue && !day.HasValue;
        if (month.HasValue && (month.Value < 1 || month.Value > 12)) return false;
        if (day.HasValue && (!month.HasValue || day.Value < 1 || day.Value > DateTime.DaysInMonth(Math.Max(1, Math.Min(9999, year.Value)), month.Value))) return false;
        return year.Value >= 1 && year.Value <= 9999;
    }

    private static void InspectProperties(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report) {
        if (format == BibliographyFormat.Nbib) {
            Check(item.Publisher, "publisher"); Check(item.PublisherPlace, "publisher-place"); Check(item.Edition, "edition"); Check(item.Url, "URL"); Check(item.CollectionTitle, "collection-title");
        } else if (format == BibliographyFormat.Ris) Check(item.CollectionTitle, "collection-title");
        void Check(string? value, string field) { if (!string.IsNullOrWhiteSpace(value)) Loss(report, item, field, "BIBCONV203", $"Field '{field}' is not represented in {format}.", BibliographyConversionAction.Omitted); }
    }

    private static void InspectNestedNativeFields(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report) {
        if (format == BibliographyFormat.CslJson) return;
        foreach (BibliographyContributor contributor in item.Contributors) foreach (BibliographyNativeField field in contributor.Name.NativeFields) Loss(report, item, "contributors." + field.Name, "BIBCONV213", $"Native name property '{field.Name}' cannot be represented in {format}.", BibliographyConversionAction.Omitted);
        foreach (BibliographyDate date in item.Dates) foreach (BibliographyNativeField field in date.NativeFields) Loss(report, item, "dates." + field.Name, "BIBCONV214", $"Native date property '{field.Name}' cannot be represented in {format}.", BibliographyConversionAction.Omitted);
    }

    private static void InspectIdentifiers(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report) {
        if (format == BibliographyFormat.CslJson) {
            foreach (IGrouping<string, BibliographyIdentifier> group in item.Identifiers.GroupBy(static identifier => identifier.Scheme, StringComparer.OrdinalIgnoreCase)) {
                if (!CodecMappings.IsCslIdentifierScheme(group.Key)) Loss(report, item, "identifiers." + group.Key, "BIBCONV225", $"Identifier scheme '{group.Key}' is not represented by the typed CSL JSON model.", BibliographyConversionAction.Omitted);
                else if (group.Count() > 1) Loss(report, item, "identifiers." + group.Key, "BIBCONV206", $"Multiple '{group.Key}' identifiers collapse into one destination value in {format}.", BibliographyConversionAction.Approximated);
            }
        }
        if (format == BibliographyFormat.Ris) {
            foreach (BibliographyIdentifier identifier in item.Identifiers.Where(static identifier => !TaggedCodec.CanRoundTripRisIdentifier(identifier)))
                Loss(report, item, "identifiers." + identifier.Scheme, "BIBCONV228", $"Identifier scheme '{identifier.Scheme}' cannot be represented unambiguously in RIS AN output.", BibliographyConversionAction.Approximated);
        }
        if (format != BibliographyFormat.EndNoteXml) return;
        foreach (BibliographyIdentifier identifier in item.Identifiers) {
            if (!string.Equals(identifier.Scheme, "ISBN", StringComparison.OrdinalIgnoreCase) && !string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase) && !string.Equals(identifier.Scheme, "DOI", StringComparison.OrdinalIgnoreCase) && !string.Equals(identifier.Scheme, "accession", StringComparison.OrdinalIgnoreCase))
                Loss(report, item, "identifiers." + identifier.Scheme, "BIBCONV204", $"Identifier scheme '{identifier.Scheme}' is not represented in EndNote XML.", BibliographyConversionAction.Omitted);
        }
    }

    private static void InspectRepeatableValues(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report) {
        if (item.Notes.Count > 1 && format != BibliographyFormat.Ris && format != BibliographyFormat.Nbib) Loss(report, item, "notes", "BIBCONV207", $"Multiple notes collapse into one destination value in {format}.", BibliographyConversionAction.Approximated);
        if (item.Keywords.Count > 1 && format == BibliographyFormat.CslJson) Loss(report, item, "keywords", "BIBCONV208", $"Multiple keywords collapse into one destination value in {format}.", BibliographyConversionAction.Approximated);
    }

    private static void InspectTextEncoding(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report) {
        foreach (KeyValuePair<string, string> text in EnumerateText(item)) {
            if ((format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib) && (text.Value.IndexOf('\r') >= 0 || text.Value.IndexOf('\n') >= 0))
                Loss(report, item, text.Key, "BIBCONV209", $"Line breaks in '{text.Key}' normalize to tagged-format continuations in {format}.", BibliographyConversionAction.Approximated);
            if (format == BibliographyFormat.EndNoteXml && HasInvalidXmlCharacters(text.Value))
                Loss(report, item, text.Key, "BIBCONV210", $"Invalid XML characters in '{text.Key}' are replaced in EndNote XML.", BibliographyConversionAction.Approximated);
            if ((format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex) && !HasBalancedBraces(text.Value))
                Loss(report, item, text.Key, "BIBCONV211", $"Unbalanced braces in '{text.Key}' are escaped for safe BibTeX output.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib) {
            foreach (BibliographyNativeField field in item.NativeFields.Where(field => field.Format == format && (field.Value.IndexOf('\r') >= 0 || field.Value.IndexOf('\n') >= 0)))
                Loss(report, item, "native." + field.Name, "BIBCONV209", $"Line breaks in native field '{field.Name}' normalize to tagged-format continuations in {format}.", BibliographyConversionAction.Approximated);
        }
    }

    private static IEnumerable<KeyValuePair<string, string>> EnumerateText(BibliographyItem item) {
        string?[] values = { item.Key, item.Title, item.ContainerTitle, item.CollectionTitle, item.Publisher, item.PublisherPlace, item.Edition, item.Volume, item.Issue, item.Pages, item.Abstract, item.Language, item.Url };
        string[] names = { "key", "title", "container-title", "collection-title", "publisher", "publisher-place", "edition", "volume", "issue", "pages", "abstract", "language", "URL" };
        for (int index = 0; index < values.Length; index++) if (!string.IsNullOrEmpty(values[index])) yield return new KeyValuePair<string, string>(names[index], values[index]!);
        foreach (BibliographyContributor contributor in item.Contributors) foreach (string? value in new[] { contributor.Name.Given, contributor.Name.Family, contributor.Name.Literal, contributor.Name.Suffix, contributor.Name.DroppingParticle, contributor.Name.NonDroppingParticle }) if (!string.IsNullOrEmpty(value)) yield return new KeyValuePair<string, string>("contributors", value!);
        foreach (BibliographyIdentifier identifier in item.Identifiers) yield return new KeyValuePair<string, string>("identifiers." + identifier.Scheme, identifier.Value);
        foreach (string value in item.Keywords) yield return new KeyValuePair<string, string>("keywords", value);
        foreach (string value in item.Notes) yield return new KeyValuePair<string, string>("notes", value);
    }

    private static bool HasBalancedBraces(string value) {
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\\' && index + 1 < value.Length) { index++; continue; }
            if (value[index] == '{') depth++;
            else if (value[index] == '}' && --depth < 0) return false;
        }
        return depth == 0;
    }

    private static bool HasInvalidXmlCharacters(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) { index++; continue; }
            if (!System.Xml.XmlConvert.IsXmlChar(value[index])) return true;
        }
        return false;
    }

    private static void Loss(BibliographyConversionReport report, BibliographyItem item, string field, string code, string message, BibliographyConversionAction action) =>
        report.Add(code, BibliographyDiagnosticSeverity.Warning, message, action, item, field);
}
