namespace OfficeIMO.Bibliography;

internal static class BibliographyWriter {
    internal static BibliographyWriteResult Write(BibliographyDocument document, BibliographyWriteOptions? options, CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new BibliographyWriteOptions();
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        BibliographyFormat format = options.Format ?? document.SourceFormat;
        var report = new BibliographyConversionReport();

        if (options.Mode == BibliographyWriterMode.Preserve && format == document.SourceFormat && !document.IsModifiedWithCancellation(cancellationToken) && document.OriginalText != null) {
            Encoding preservedEncoding = document.OriginalBytes == null ? ResolvePreservedEncoding(document.OriginalText, format, options.Encoding, cancellationToken) : options.Encoding;
            if (document.OriginalBytes == null) InspectEncoding(document.OriginalText, preservedEncoding, report, cancellationToken);
            if (options.RequireNoLoss) report.RequireNoLoss();
            byte[] bytes = document.OriginalBytes != null ? BibliographyEncoding.CloneBytes(document.OriginalBytes, cancellationToken) : BibliographyEncoding.Encode(document.OriginalText, preservedEncoding, cancellationToken);
            return new BibliographyWriteResult(document.OriginalText, bytes, format, true, report);
        }

        BibliographyConversionInspector.Inspect(document, format, report, cancellationToken);
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
        InspectEncoding(content, options.Encoding, report, cancellationToken);
        if (options.RequireNoLoss) report.RequireNoLoss();
        return new BibliographyWriteResult(content, BibliographyEncoding.Encode(content, options.Encoding, cancellationToken), format, false, report);
    }

    private static void InspectEncoding(string value, Encoding encoding, BibliographyConversionReport report, CancellationToken cancellationToken) {
        if (!BibliographyEncoding.CanEncode(value, encoding, cancellationToken)) {
            report.Add("BIBCONV220", BibliographyDiagnosticSeverity.Warning, $"The selected {encoding.WebName} encoding cannot represent all output characters without replacement.", BibliographyConversionAction.Approximated, field: "encoding");
        }
    }

    private static Encoding ResolvePreservedEncoding(string source, BibliographyFormat format, Encoding fallback, CancellationToken cancellationToken) {
        return format == BibliographyFormat.EndNoteXml ? BibliographyEncoding.ResolveXmlDeclaration(source, fallback, cancellationToken) : fallback;
    }
}

internal static class BibliographyConversionInspector {
    internal static void Inspect(BibliographyDocument document, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        foreach (BibliographyDiagnostic diagnostic in Cancellable(document.Diagnostics, cancellationToken).Where(IsRecoveryLossDiagnostic)) {
            report.Add("BIBCONV222", BibliographyDiagnosticSeverity.Warning, $"Canonical output is based on partially recovered source after parser diagnostic {diagnostic.Code}; unrecovered source content may be omitted.", BibliographyConversionAction.Omitted, field: diagnostic.Field);
        }
        InspectKeys(document, format, report, cancellationToken);
        InspectDocumentStructure(document, format, report, cancellationToken);
        foreach (BibliographyItem item in Cancellable(document.Items, cancellationToken)) {
            InspectType(item, document.SourceFormat, format, report, cancellationToken); InspectContributors(item, format, report, cancellationToken); InspectDates(item, format, report, cancellationToken); InspectNestedNativeFields(item, format, report, cancellationToken); InspectProperties(item, format, report, cancellationToken); InspectIdentifiers(item, format, report, cancellationToken); InspectRepeatableValues(item, format, report); InspectTextEncoding(item, format, report, cancellationToken); InspectNativeStructure(item, format, report, cancellationToken);
        }
    }

    private static bool IsRecoveryLossDiagnostic(BibliographyDiagnostic diagnostic) =>
        diagnostic.Severity == BibliographyDiagnosticSeverity.Error || diagnostic.Code == "BIBBIB001" || diagnostic.Code == "BIBTAG001" || diagnostic.Code == "BIBTAG004" || diagnostic.Code == "BIBCSL003" || diagnostic.Code == "BIBEND004" || diagnostic.Code == "BIBEND005";

    private static void InspectKeys(BibliographyDocument document, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        foreach (BibliographyItem item in Cancellable(document.Items, cancellationToken).Where(item => string.IsNullOrWhiteSpace(item.Key) && !(format == BibliographyFormat.CslJson && string.IsNullOrEmpty(item.Key) && CslJsonCodec.HasNativeProperty(item, "id", cancellationToken))))
            Loss(report, item, "key", "BIBCONV215", $"A missing citation key is replaced with a deterministic generated identifier in {format}.", BibliographyConversionAction.Approximated);
        StringComparer keyComparer = format == BibliographyFormat.CslJson ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        foreach (IGrouping<string, BibliographyItem> duplicate in Cancellable(document.Items, cancellationToken).Where(static item => !string.IsNullOrWhiteSpace(item.Key)).GroupBy(item => CodecMappings.NormalizeOutputKey(item.Key, format), keyComparer).Where(group => Cancellable(group, cancellationToken).Skip(1).Any())) {
            foreach (BibliographyItem item in Cancellable(duplicate, cancellationToken)) Loss(report, item, "key", "BIBCONV216", $"Duplicate citation key '{duplicate.Key}' is not unique in {format} output.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex) {
            foreach (BibliographyItem item in Cancellable(document.Items, cancellationToken).Where(static item => !string.IsNullOrWhiteSpace(item.Key) && item.Key.Any(character => !BibCodec.IsSafeKeyCharacter(character))))
                Loss(report, item, "key", "BIBCONV217", "The citation key contains characters that must be normalized for BibTeX output.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.Nbib) {
            foreach (BibliographyItem item in Cancellable(document.Items, cancellationToken)) {
                BibliographyIdentifier[] pmids = Cancellable(item.Identifiers, cancellationToken).Where(static identifier => string.Equals(identifier.Scheme, "PMID", StringComparison.OrdinalIgnoreCase)).ToArray();
                string? pmid = pmids.FirstOrDefault()?.Value;
                if (string.IsNullOrWhiteSpace(pmid)) Loss(report, item, "identifiers.PMID", "BIBCONV223", "NBIB output uses the citation key as a PMID because the item has no PMID identifier.", BibliographyConversionAction.Approximated);
                else if (!string.Equals(item.Key, pmid, StringComparison.Ordinal)) Loss(report, item, "key", "BIBCONV224", "NBIB represents PMID as its record identifier, so the distinct citation key is omitted.", BibliographyConversionAction.Omitted);
                if (pmids.Length > 1) Loss(report, item, "identifiers.PMID", "BIBCONV227", "NBIB represents one PMID as record identity, so additional PMID identifiers are omitted.", BibliographyConversionAction.Omitted);
            }
        }
    }

    private static void InspectType(BibliographyItem item, BibliographyFormat sourceFormat, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        if (format == BibliographyFormat.CslJson && CslJsonCodec.UsesNativeType(sourceFormat, item) && CslJsonCodec.ContainsInvalidUtf16(item.NativeType!, cancellationToken))
            Loss(report, item, "type", "BIBCONV250", "Invalid UTF-16 in the native item type is replaced during CSL JSON serialization.", BibliographyConversionAction.Approximated);
        if (format == BibliographyFormat.EndNoteXml && EndNoteXmlCodec.CanPreserveNativeType(sourceFormat, item) && HasInvalidXmlCharacters(item.NativeType!, cancellationToken))
            Loss(report, item, "type", "BIBCONV210", "Invalid XML characters in the native item type are replaced in EndNote XML.", BibliographyConversionAction.Approximated);
        bool exact;
        switch (format) {
            case BibliographyFormat.CslJson:
                exact = CslJsonCodec.PreservesNativeType(sourceFormat, item) && (CslJsonCodec.CanRoundTripType(sourceFormat, item) || IsExactCslType(item.Type)) ||
                    item.Type == BibliographyItemType.Unknown && item.NativeType == null && CslJsonCodec.HasNativeProperty(item, "type", cancellationToken);
                break;
            case BibliographyFormat.BibTex: case BibliographyFormat.BibLatex:
                bool hasNativeBibType = (sourceFormat == BibliographyFormat.BibTex || sourceFormat == BibliographyFormat.BibLatex) && !string.IsNullOrWhiteSpace(item.NativeType);
                exact = hasNativeBibType && CodecMappings.ParseType(item.NativeType) == item.Type
                    ? BibCodec.CanPreserveNativeType(sourceFormat, format, item)
                    : BibCodec.CanRoundTripType(item.Type, format);
                break;
            case BibliographyFormat.Ris:
                exact = sourceFormat == BibliographyFormat.Ris && item.NativeType != null
                    ? TaggedCodec.CanPreserveNativeType(sourceFormat, item) ||
                      item.Type == BibliographyItemType.Unknown && TaggedCodec.CanPreserveUnknownRisType(item.NativeType)
                    : TaggedCodec.CanRoundTripRisType(item.Type);
                break;
            case BibliographyFormat.Nbib:
                exact = TaggedCodec.CanRoundTripNbibType(item.Type) ||
                    item.Type != BibliographyItemType.Unknown && sourceFormat == BibliographyFormat.Nbib &&
                    Cancellable(item.NativeFields, cancellationToken).Any(field => field.Format == BibliographyFormat.Nbib && string.Equals(field.Name, "PT", StringComparison.OrdinalIgnoreCase) && CodecMappings.ParseType(field.Value) == item.Type);
                break;
            case BibliographyFormat.EndNoteXml:
                exact = EndNoteXmlCodec.CanPreserveNativeType(sourceFormat, item) ||
                    (item.Type == BibliographyItemType.ArticleJournal || item.Type == BibliographyItemType.Book || item.Type == BibliographyItemType.Chapter || item.Type == BibliographyItemType.PaperConference || item.Type == BibliographyItemType.Report || item.Type == BibliographyItemType.Thesis || item.Type == BibliographyItemType.WebPage || item.Type == BibliographyItemType.Patent || item.Type == BibliographyItemType.Document) &&
                    (sourceFormat != BibliographyFormat.EndNoteXml || string.IsNullOrWhiteSpace(item.NativeType) || EndNoteXmlCodec.CanPreserveNativeType(sourceFormat, item));
                break;
            default: exact = false; break;
        }
        if (!exact) Loss(report, item, "type", "BIBCONV200", $"Item type '{item.Type}' is written using a broader {format} type.", BibliographyConversionAction.Approximated);
    }

    private static bool IsExactCslType(BibliographyItemType type) {
        switch (type) {
            case BibliographyItemType.Article: case BibliographyItemType.ArticleJournal: case BibliographyItemType.ArticleMagazine: case BibliographyItemType.ArticleNewspaper:
            case BibliographyItemType.Book: case BibliographyItemType.Chapter: case BibliographyItemType.PaperConference: case BibliographyItemType.Report:
            case BibliographyItemType.Thesis: case BibliographyItemType.WebPage: case BibliographyItemType.Dataset: case BibliographyItemType.Software:
            case BibliographyItemType.Patent: case BibliographyItemType.LegalCase: case BibliographyItemType.Manuscript: case BibliographyItemType.PersonalCommunication:
            case BibliographyItemType.Document: return true;
            default: return false;
        }
    }

    private static void InspectContributors(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        foreach (BibliographyContributorRole role in Cancellable(item.Contributors, cancellationToken).Select(static value => value.Role).Distinct()) {
            bool exact;
            switch (format) {
                case BibliographyFormat.BibTex: exact = role == BibliographyContributorRole.Author || role == BibliographyContributorRole.Editor; break;
                case BibliographyFormat.BibLatex: exact = role == BibliographyContributorRole.Author || role == BibliographyContributorRole.Editor || role == BibliographyContributorRole.Translator; break;
                case BibliographyFormat.CslJson: exact = role == BibliographyContributorRole.Author || role == BibliographyContributorRole.Editor || role == BibliographyContributorRole.Translator || role == BibliographyContributorRole.Recipient || role == BibliographyContributorRole.Interviewer || role == BibliographyContributorRole.Composer || role == BibliographyContributorRole.CollectionEditor; break;
                case BibliographyFormat.Ris: exact = role == BibliographyContributorRole.Author || role == BibliographyContributorRole.Editor; break;
                case BibliographyFormat.Nbib: exact = role == BibliographyContributorRole.Author; break;
                case BibliographyFormat.EndNoteXml: exact = role == BibliographyContributorRole.Author || role == BibliographyContributorRole.Editor || role == BibliographyContributorRole.CollectionEditor || role == BibliographyContributorRole.Translator; break;
                default: exact = false; break;
            }
            if (!exact) Loss(report, item, "contributors." + role, "BIBCONV201", $"Contributor role '{role}' is not represented exactly in {format}.", format == BibliographyFormat.EndNoteXml ? BibliographyConversionAction.Approximated : BibliographyConversionAction.Omitted);
        }
        if (format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex) {
            foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken).Where(static contributor => !BibCodec.CanRoundTripStructuredName(contributor.Name)))
                Loss(report, item, "contributors", "BIBCONV226", "A structured contributor name cannot be reopened exactly through BibTeX name syntax.", BibliographyConversionAction.Approximated);
        } else if (format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib || format == BibliographyFormat.EndNoteXml) {
            foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken).Where(static contributor => !string.IsNullOrWhiteSpace(contributor.Name.DroppingParticle) || !string.IsNullOrWhiteSpace(contributor.Name.NonDroppingParticle)))
                Loss(report, item, "contributors", "BIBCONV229", $"Structured contributor particles are flattened in {format} output and cannot be reopened exactly.", BibliographyConversionAction.Approximated);
            foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken).Where(static contributor => !string.IsNullOrWhiteSpace(contributor.Name.Literal) && (!string.IsNullOrWhiteSpace(contributor.Name.Given) || !string.IsNullOrWhiteSpace(contributor.Name.Family) || !string.IsNullOrWhiteSpace(contributor.Name.Suffix))))
                Loss(report, item, "contributors", "BIBCONV231", $"A literal contributor also has personal-name components that are omitted in {format} output.", BibliographyConversionAction.Omitted);
            foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken).Where(static contributor => ContainsComma(contributor.Name.Given) || ContainsComma(contributor.Name.Family) || ContainsComma(contributor.Name.Suffix)))
                Loss(report, item, "contributors", "BIBCONV236", $"A structured contributor name contains a comma that is indistinguishable from {format} name-component separators.", BibliographyConversionAction.Approximated);
            foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken).Where(static contributor => HasSurroundingWhitespace(contributor.Name.Given) || HasSurroundingWhitespace(contributor.Name.Family) || HasSurroundingWhitespace(contributor.Name.Suffix) || HasSurroundingWhitespace(contributor.Name.DroppingParticle) || HasSurroundingWhitespace(contributor.Name.NonDroppingParticle) || HasLeadingWhitespace(contributor.Name.Literal)))
                Loss(report, item, "contributors", "BIBCONV243", $"A contributor name contains whitespace that is trimmed by {format} name parsing.", BibliographyConversionAction.Approximated);
        }
        if (format != BibliographyFormat.CslJson) {
            foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken).Where(static contributor => HasEmptyNameComponent(contributor.Name)))
                Loss(report, item, "contributors", "BIBCONV244", $"An empty or entirely unset contributor name cannot reopen exactly in {format}.", BibliographyConversionAction.Approximated);
        }
        if (ReordersContributors(item, format, cancellationToken))
            Loss(report, item, "contributors", "BIBCONV230", $"Contributor source order is regrouped by {format} output and cannot be reopened exactly.", BibliographyConversionAction.Approximated);
    }

    private static bool ReordersContributors(BibliographyItem item, BibliographyFormat format, CancellationToken cancellationToken) {
        BibliographyContributor[] source;
        BibliographyContributor[] output;
        switch (format) {
            case BibliographyFormat.BibTex: case BibliographyFormat.BibLatex:
                BibliographyContributorRole[] bibRoles = { BibliographyContributorRole.Author, BibliographyContributorRole.Editor, BibliographyContributorRole.Translator };
                source = Cancellable(item.Contributors, cancellationToken).Where(contributor => bibRoles.Contains(contributor.Role)).ToArray();
                output = bibRoles.SelectMany(role => Cancellable(source, cancellationToken).Where(contributor => contributor.Role == role)).ToArray();
                break;
            case BibliographyFormat.CslJson:
                BibliographyContributorRole[] cslRoles = { BibliographyContributorRole.Author, BibliographyContributorRole.Editor, BibliographyContributorRole.Translator, BibliographyContributorRole.Recipient, BibliographyContributorRole.Interviewer, BibliographyContributorRole.Composer, BibliographyContributorRole.CollectionEditor };
                source = Cancellable(item.Contributors, cancellationToken).Where(contributor => cslRoles.Contains(contributor.Role)).ToArray();
                output = cslRoles.SelectMany(role => Cancellable(source, cancellationToken).Where(contributor => contributor.Role == role)).ToArray();
                break;
            case BibliographyFormat.EndNoteXml:
                BibliographyContributorRole[] endNoteRoles = { BibliographyContributorRole.Author, BibliographyContributorRole.Editor, BibliographyContributorRole.CollectionEditor, BibliographyContributorRole.Translator };
                source = Cancellable(item.Contributors, cancellationToken).Where(contributor => endNoteRoles.Contains(contributor.Role)).ToArray();
                output = Cancellable(source, cancellationToken).GroupBy(static contributor => contributor.Role).SelectMany(group => Cancellable(group, cancellationToken)).ToArray();
                break;
            case BibliographyFormat.Nbib:
                source = Cancellable(item.Contributors, cancellationToken).Where(static contributor => contributor.Role == BibliographyContributorRole.Author).ToArray();
                output = Cancellable(source, cancellationToken).Where(static contributor => string.IsNullOrWhiteSpace(contributor.Name.Literal)).Concat(Cancellable(source, cancellationToken).Where(static contributor => !string.IsNullOrWhiteSpace(contributor.Name.Literal))).ToArray();
                break;
            default:
                return false;
        }
        return !source.SequenceEqual(output);
    }

    private static bool ContainsComma(string? value) => value?.IndexOf(',') >= 0;
    private static bool HasSurroundingWhitespace(string? value) => value != null && !string.Equals(value, value.Trim(), StringComparison.Ordinal);
    private static bool HasLeadingWhitespace(string? value) => !string.IsNullOrEmpty(value) && char.IsWhiteSpace(value![0]);
    private static bool HasEmptyNameComponent(BibliographyName name) =>
        name.Given == null && name.Family == null && name.Literal == null && name.Suffix == null && name.DroppingParticle == null && name.NonDroppingParticle == null ||
        name.Literal == null && name.Family == null ||
        name.Given is { Length: 0 } || name.Family is { Length: 0 } || name.Literal is { Length: 0 } || name.Suffix is { Length: 0 } ||
        name.DroppingParticle is { Length: 0 } || name.NonDroppingParticle is { Length: 0 };

    private static void InspectDocumentStructure(BibliographyDocument document, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        if (format != BibliographyFormat.EndNoteXml) return;
        string rootElementName = document.EndNoteRootElementName ?? (document.EndNoteRecordsRoot ? "records" : "xml");
        string recordsElementName = document.EndNoteRecordsElementName ?? "records";
        if (EndNoteXmlCodec.HasDuplicateDocumentAttributes(document, rootElementName, cancellationToken))
            report.Add("BIBCONV248", BibliographyDiagnosticSeverity.Warning, $"Additional EndNote XML attribute metadata for '{rootElementName}' is omitted by canonical output.", BibliographyConversionAction.Omitted, field: rootElementName);
        if (!string.Equals(recordsElementName, rootElementName, StringComparison.OrdinalIgnoreCase) && EndNoteXmlCodec.CoalescesRecordsContainerMetadata(document, cancellationToken))
            report.Add("BIBCONV238", BibliographyDiagnosticSeverity.Warning, "Separate EndNote XML records-container metadata is coalesced into one canonical records container.", BibliographyConversionAction.Approximated, field: "records");
    }

    private static void InspectDates(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        foreach (BibliographyDateRole role in Cancellable(item.Dates, cancellationToken).Select(static value => value.Role).Distinct()) {
            bool exact = format == BibliographyFormat.CslJson ? role == BibliographyDateRole.Issued || role == BibliographyDateRole.Accessed || role == BibliographyDateRole.Submitted || role == BibliographyDateRole.Original || role == BibliographyDateRole.Event
                : (format == BibliographyFormat.BibLatex || format == BibliographyFormat.Ris) ? role == BibliographyDateRole.Issued || role == BibliographyDateRole.Accessed
                : format == BibliographyFormat.BibTex ? role == BibliographyDateRole.Issued
                : role == BibliographyDateRole.Issued;
            if (!exact) Loss(report, item, "dates." + role, "BIBCONV202", $"Date role '{role}' is not represented in {format}.", BibliographyConversionAction.Omitted);
            if (Cancellable(item.Dates, cancellationToken).Count(date => date.Role == role) > 1) Loss(report, item, "dates." + role, "BIBCONV205", $"Multiple '{role}' dates collapse to the first value in {format}.", BibliographyConversionAction.Approximated);
        }
        BibliographyDate? issued = item.GetDate(BibliographyDateRole.Issued);
        if (format == BibliographyFormat.BibTex && issued?.Day != null) Loss(report, item, "dates.Issued.day", "BIBCONV212", "Classic BibTeX output omits issued-day precision.", BibliographyConversionAction.Omitted);
        if (format == BibliographyFormat.BibTex && issued != null && !issued.Year.HasValue && !issued.Month.HasValue && !issued.Day.HasValue && issued.Literal == null)
            Loss(report, item, "dates.Issued", "BIBCONV241", "Classic BibTeX output omits an issued date with no representable year, month, day, or literal value.", BibliographyConversionAction.Omitted);
        foreach (BibliographyDate date in Cancellable(item.Dates, cancellationToken)) {
            if ((format == BibliographyFormat.BibLatex || format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib || format == BibliographyFormat.EndNoteXml) &&
                !date.Year.HasValue && !date.Month.HasValue && !date.Day.HasValue && !date.EndYear.HasValue && !date.EndMonth.HasValue && !date.EndDay.HasValue && date.Literal == null)
                Loss(report, item, "dates." + date.Role, "BIBCONV242", $"A null-valued empty date reopens with an empty literal in {format}.", BibliographyConversionAction.Approximated);
            bool classicBibMonthOnly = format == BibliographyFormat.BibTex && date.Role == BibliographyDateRole.Issued && !date.Year.HasValue && date.Month is >= 1 and <= 12 && !date.Day.HasValue;
            bool cslOmitsNumericParts = format == BibliographyFormat.CslJson &&
                (!date.Year.HasValue && (date.Month.HasValue || date.Day.HasValue || date.EndYear.HasValue || date.EndMonth.HasValue || date.EndDay.HasValue) ||
                 date.Day.HasValue && !date.Month.HasValue ||
                 !date.EndYear.HasValue && (date.EndMonth.HasValue || date.EndDay.HasValue) ||
                 date.EndDay.HasValue && !date.EndMonth.HasValue);
            bool cslHasOutOfRangeParts = format == BibliographyFormat.CslJson &&
                (date.Month is < 1 or > 12 || date.Day is < 1 or > 31 || date.EndMonth is < 1 or > 12 || date.EndDay is < 1 or > 31);
            if (cslOmitsNumericParts || cslHasOutOfRangeParts || format != BibliographyFormat.CslJson && ((!classicBibMonthOnly && !IsValidDate(date.Year, date.Month, date.Day)) || !IsValidDate(date.EndYear, date.EndMonth, date.EndDay) || date.EndYear.HasValue && !date.Year.HasValue))
                Loss(report, item, "dates." + date.Role, "BIBCONV218", "A date contains an invalid or incomplete numeric component sequence.", BibliographyConversionAction.Approximated);
            if (date.EndYear.HasValue && !CanRoundTripDateRange(format, date.Role))
                Loss(report, item, "dates." + date.Role + ".end", "BIBCONV219", $"Date ranges are not represented exactly in {format}.", BibliographyConversionAction.Approximated);
            if (format != BibliographyFormat.CslJson && date.Year.HasValue && date.Literal != null)
                Loss(report, item, "dates." + date.Role + ".literal", "BIBCONV221", $"The literal date value is not represented alongside numeric date parts in {format}.", BibliographyConversionAction.Omitted);
            if ((format == BibliographyFormat.BibLatex || format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib || format == BibliographyFormat.EndNoteXml) &&
                !date.Year.HasValue && !string.IsNullOrWhiteSpace(date.Literal) && CodecMappings.IsStructuredDateText(date.Literal!))
                Loss(report, item, "dates." + date.Role + ".literal", "BIBCONV240", $"A parseable literal date is reopened as structured numeric parts in {format}.", BibliographyConversionAction.Approximated);
            if ((format == BibliographyFormat.BibLatex || format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib || format == BibliographyFormat.EndNoteXml) &&
                LiteralDateWhitespaceIsNormalized(date.Literal, format))
                Loss(report, item, "dates." + date.Role + ".literal", "BIBCONV246", $"Literal-date surrounding whitespace is trimmed by {format} date parsing.", BibliographyConversionAction.Approximated);
        }
    }

    private static bool CanRoundTripDateRange(BibliographyFormat format, BibliographyDateRole role) {
        if (format == BibliographyFormat.CslJson)
            return role == BibliographyDateRole.Issued || role == BibliographyDateRole.Accessed || role == BibliographyDateRole.Submitted || role == BibliographyDateRole.Original || role == BibliographyDateRole.Event;
        if (format == BibliographyFormat.BibLatex || format == BibliographyFormat.Ris)
            return role == BibliographyDateRole.Issued || role == BibliographyDateRole.Accessed;
        return (format == BibliographyFormat.Nbib || format == BibliographyFormat.EndNoteXml) && role == BibliographyDateRole.Issued;
    }

    private static bool LiteralDateWhitespaceIsNormalized(string? value, BibliographyFormat format) {
        if (value == null || string.Equals(value, value.Trim(), StringComparison.Ordinal)) return false;
        return format != BibliographyFormat.EndNoteXml || value.Trim().Length > 0;
    }

    private static bool IsValidDate(int? year, int? month, int? day) {
        if (!year.HasValue) return !month.HasValue && !day.HasValue;
        if (month.HasValue && (month.Value < 1 || month.Value > 12)) return false;
        if (day.HasValue && (!month.HasValue || day.Value < 1 || day.Value > 31)) return false;
        return year.Value >= 1;
    }

    internal static void InspectProperties(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        if (format == BibliographyFormat.Nbib) {
            Check(item.Publisher, "publisher"); Check(item.PublisherPlace, "publisher-place"); Check(item.Edition, "edition"); Check(item.Url, "URL"); Check(item.CollectionTitle, "collection-title");
        } else if (format == BibliographyFormat.Ris) Check(item.CollectionTitle, "collection-title");
        else if (format == BibliographyFormat.EndNoteXml && item.Url != null && item.Url.Length == 0 && Cancellable(item.NativeFields, cancellationToken).Any(static field => field.Format == BibliographyFormat.EndNoteXml && string.Equals(field.Name, "url", StringComparison.OrdinalIgnoreCase)))
            Loss(report, item, "URL", "BIBCONV237", "An empty primary EndNote URL with additional URL roles reopens as a missing primary URL.", BibliographyConversionAction.Approximated);
        void Check(string? value, string field) { if (value != null) Loss(report, item, field, "BIBCONV203", $"Field '{field}' is not represented in {format}.", BibliographyConversionAction.Omitted); }
    }

    private static void InspectNestedNativeFields(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        if (format == BibliographyFormat.CslJson) return;
        foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken)) foreach (BibliographyNativeField field in Cancellable(contributor.Name.NativeFields, cancellationToken)) Loss(report, item, "contributors." + field.Name, "BIBCONV213", $"Native name property '{field.Name}' cannot be represented in {format}.", BibliographyConversionAction.Omitted);
        foreach (BibliographyDate date in Cancellable(item.Dates, cancellationToken)) foreach (BibliographyNativeField field in Cancellable(date.NativeFields, cancellationToken)) {
            if (format == BibliographyFormat.EndNoteXml && (EndNoteXmlCodec.CanPreserveNativeDateField(date, field, cancellationToken) || EndNoteXmlCodec.CanPreserveNativePublicationDateField(date, field, cancellationToken))) continue;
            Loss(report, item, "dates." + field.Name, "BIBCONV214", $"Native date property '{field.Name}' cannot be represented in {format}.", BibliographyConversionAction.Omitted);
        }
    }

    private static void InspectIdentifiers(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        if (format == BibliographyFormat.CslJson) {
            foreach (IGrouping<string, BibliographyIdentifier> group in Cancellable(item.Identifiers, cancellationToken).GroupBy(static identifier => identifier.Scheme, StringComparer.OrdinalIgnoreCase)) {
                if (!CodecMappings.IsCslIdentifierScheme(group.Key)) Loss(report, item, "identifiers." + group.Key, "BIBCONV225", $"Identifier scheme '{group.Key}' is not represented by the typed CSL JSON model.", BibliographyConversionAction.Omitted);
                else if (group.Count() > 1) Loss(report, item, "identifiers." + group.Key, "BIBCONV206", $"Multiple '{group.Key}' identifiers collapse into one destination value in {format}.", BibliographyConversionAction.Approximated);
            }
            foreach (BibliographyIdentifier identifier in Cancellable(item.Identifiers, cancellationToken).Where(static identifier => CodecMappings.IsCslIdentifierScheme(identifier.Scheme) && !CodecMappings.IsCanonicalCslIdentifierScheme(identifier.Scheme)))
                Loss(report, item, "identifiers." + identifier.Scheme, "BIBCONV245", $"Identifier scheme '{identifier.Scheme}' is normalized to uppercase CSL JSON property spelling.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.Ris) {
            foreach (BibliographyIdentifier identifier in Cancellable(item.Identifiers, cancellationToken).Where(static identifier => !TaggedCodec.CanRoundTripRisIdentifier(identifier)))
                Loss(report, item, "identifiers." + identifier.Scheme, "BIBCONV228", $"Identifier scheme '{identifier.Scheme}' cannot be represented unambiguously in RIS AN output.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.Nbib) {
            foreach (BibliographyIdentifier identifier in Cancellable(item.Identifiers, cancellationToken).Where(static identifier => !TaggedCodec.CanRoundTripNbibIdentifier(identifier)))
                Loss(report, item, "identifiers." + identifier.Scheme, "BIBCONV232", $"Identifier scheme '{identifier.Scheme}' cannot be represented unambiguously in NBIB output.", BibliographyConversionAction.Omitted);
        }
        if (format != BibliographyFormat.EndNoteXml) return;
        foreach (BibliographyIdentifier identifier in Cancellable(item.Identifiers, cancellationToken)) {
            bool exactSerial = string.Equals(identifier.Scheme, "ISBN", StringComparison.OrdinalIgnoreCase) || string.Equals(identifier.Scheme, "ISSN", StringComparison.OrdinalIgnoreCase);
            bool exactDoi = string.Equals(identifier.Scheme, "DOI", StringComparison.Ordinal);
            bool exactAccession = string.Equals(identifier.Scheme, "accession", StringComparison.Ordinal);
            if (!exactSerial && !exactDoi && !exactAccession)
                Loss(report, item, "identifiers." + identifier.Scheme, "BIBCONV204", $"Identifier scheme '{identifier.Scheme}' is not represented in EndNote XML.", BibliographyConversionAction.Omitted);
        }
    }

    private static void InspectRepeatableValues(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report) {
        if (item.Notes.Count > 1 && format != BibliographyFormat.Ris && format != BibliographyFormat.Nbib) Loss(report, item, "notes", "BIBCONV207", $"Multiple notes collapse into one destination value in {format}.", BibliographyConversionAction.Approximated);
        if (item.Keywords.Count > 1 && format == BibliographyFormat.CslJson) Loss(report, item, "keywords", "BIBCONV208", $"Multiple keywords collapse into one destination value in {format}.", BibliographyConversionAction.Approximated);
    }

    private static void InspectTextEncoding(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        foreach (KeyValuePair<string, string> text in EnumerateText(item, cancellationToken)) {
            if ((format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib) && (text.Value.IndexOf('\r') >= 0 || text.Value.IndexOf('\n') >= 0))
                Loss(report, item, text.Key, "BIBCONV209", $"Line breaks in '{text.Key}' normalize to tagged-format continuations in {format}.", BibliographyConversionAction.Approximated);
            if ((format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib) && text.Value.Length > 0 && char.IsWhiteSpace(text.Value[0]) && !IsProtectedRisIdentifierWhitespace(item, text, format, cancellationToken))
                Loss(report, item, text.Key, "BIBCONV239", $"Leading whitespace in '{text.Key}' is normalized by {format} tagged-value parsing.", BibliographyConversionAction.Approximated);
            if (format == BibliographyFormat.CslJson && CslJsonCodec.ContainsInvalidUtf16(text.Value, cancellationToken))
                Loss(report, item, text.Key, "BIBCONV250", $"Invalid UTF-16 in '{text.Key}' is replaced during CSL JSON serialization.", BibliographyConversionAction.Approximated);
            if (format == BibliographyFormat.EndNoteXml && HasInvalidXmlCharacters(text.Value, cancellationToken))
                Loss(report, item, text.Key, "BIBCONV210", $"Invalid XML characters in '{text.Key}' are replaced in EndNote XML.", BibliographyConversionAction.Approximated);
            if (format == BibliographyFormat.EndNoteXml && text.Value.IndexOf('\r') >= 0)
                Loss(report, item, text.Key, "BIBCONV235", $"Carriage returns in '{text.Key}' normalize to line feeds in EndNote XML.", BibliographyConversionAction.Approximated);
            if ((format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex) && !HasBalancedBraces(text.Value, cancellationToken))
                Loss(report, item, text.Key, "BIBCONV211", $"Unbalanced braces in '{text.Key}' are escaped for safe BibTeX output.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.Ris || format == BibliographyFormat.Nbib) {
            foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken).Where(field => field.Format == format && (field.Value.IndexOf('\r') >= 0 || field.Value.IndexOf('\n') >= 0)))
                Loss(report, item, "native." + field.Name, "BIBCONV209", $"Line breaks in native field '{field.Name}' normalize to tagged-format continuations in {format}.", BibliographyConversionAction.Approximated);
            foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken).Where(field => field.Format == format && field.Value.Length > 0 && char.IsWhiteSpace(field.Value[0])))
                Loss(report, item, "native." + field.Name, "BIBCONV239", $"Leading whitespace in native field '{field.Name}' is normalized by {format} tagged-value parsing.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.EndNoteXml) {
            foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken).Where(static field => field.Format == BibliographyFormat.EndNoteXml && field.Value.IndexOf('\r') >= 0))
                Loss(report, item, "native." + field.Name, "BIBCONV235", $"Carriage returns in native field '{field.Name}' normalize to line feeds in EndNote XML.", BibliographyConversionAction.Approximated);
        }
        if (format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex) {
            foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken).Where(field => (field.Format == BibliographyFormat.BibTex || field.Format == BibliographyFormat.BibLatex) && !HasBalancedBraces(field.Value, cancellationToken)))
                Loss(report, item, "native." + field.Name, "BIBCONV233", $"Unbalanced braces in native field '{field.Name}' are escaped for safe BibTeX output.", BibliographyConversionAction.Approximated);
        }
    }

    private static void InspectNativeStructure(BibliographyItem item, BibliographyFormat format, BibliographyConversionReport report, CancellationToken cancellationToken) {
        foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken)) InspectRaw(field, "native." + field.Name);
        foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken))
            foreach (BibliographyNativeField field in Cancellable(contributor.Name.NativeFields, cancellationToken)) InspectRaw(field, "contributors." + field.Name);
        foreach (BibliographyDate date in Cancellable(item.Dates, cancellationToken))
            foreach (BibliographyNativeField field in Cancellable(date.NativeFields, cancellationToken)) InspectRaw(field, "dates." + field.Name);
        if (format == BibliographyFormat.EndNoteXml) {
            foreach (BibliographyNativeField field in Cancellable(item.NativeFields, cancellationToken).Where(EndNoteXmlCodec.EditedNativeFieldFlattensStructure))
                Loss(report, item, "native." + field.Name, "BIBCONV234", $"Editing native EndNote field '{field.Name}' flattens its retained XML child structure.", BibliographyConversionAction.Approximated);
        }

        void InspectRaw(BibliographyNativeField field, string path) {
            if (field.Format == format && field.HasInconsistentRawValue)
                Loss(report, item, path, "BIBCONV247", $"Native {format} field '{field.Name}' has a raw representation that does not match its decoded value or field name.", BibliographyConversionAction.Approximated);
            if (format == BibliographyFormat.CslJson && field.Format == BibliographyFormat.CslJson &&
                (CslJsonCodec.ContainsInvalidUtf16(field.Name, cancellationToken) || CslJsonCodec.ContainsInvalidUtf16(field.Value, cancellationToken)))
                Loss(report, item, path, "BIBCONV250", $"Invalid UTF-16 in native CSL JSON field '{field.Name}' is replaced during serialization.", BibliographyConversionAction.Approximated);
        }
    }

    private static bool IsProtectedRisIdentifierWhitespace(BibliographyItem item, KeyValuePair<string, string> text, BibliographyFormat format, CancellationToken cancellationToken) {
        if (format != BibliographyFormat.Ris || !text.Key.StartsWith("identifiers.", StringComparison.Ordinal)) return false;
        return Cancellable(item.Identifiers, cancellationToken).Any(identifier =>
            string.Equals(text.Key, "identifiers." + identifier.Scheme, StringComparison.Ordinal) &&
            string.Equals(text.Value, identifier.Value, StringComparison.Ordinal) &&
            TaggedCodec.ProtectsLeadingWhitespaceInRisIdentifier(identifier));
    }

    private static IEnumerable<KeyValuePair<string, string>> EnumerateText(BibliographyItem item, CancellationToken cancellationToken) {
        string?[] values = { item.Key, item.Title, item.ContainerTitle, item.CollectionTitle, item.Publisher, item.PublisherPlace, item.Edition, item.Volume, item.Issue, item.Pages, item.Abstract, item.Language, item.Url };
        string[] names = { "key", "title", "container-title", "collection-title", "publisher", "publisher-place", "edition", "volume", "issue", "pages", "abstract", "language", "URL" };
        for (int index = 0; index < values.Length; index++) if (!string.IsNullOrEmpty(values[index])) yield return new KeyValuePair<string, string>(names[index], values[index]!);
        foreach (BibliographyContributor contributor in Cancellable(item.Contributors, cancellationToken)) foreach (string? value in new[] { contributor.Name.Given, contributor.Name.Family, contributor.Name.Literal, contributor.Name.Suffix, contributor.Name.DroppingParticle, contributor.Name.NonDroppingParticle }) if (!string.IsNullOrEmpty(value)) yield return new KeyValuePair<string, string>("contributors", value!);
        foreach (BibliographyIdentifier identifier in Cancellable(item.Identifiers, cancellationToken)) yield return new KeyValuePair<string, string>("identifiers." + identifier.Scheme, identifier.Value);
        foreach (BibliographyDate date in Cancellable(item.Dates, cancellationToken)) if (!string.IsNullOrEmpty(date.Literal)) yield return new KeyValuePair<string, string>("dates." + date.Role + ".literal", date.Literal!);
        foreach (string value in Cancellable(item.Keywords, cancellationToken)) yield return new KeyValuePair<string, string>("keywords", value);
        foreach (string value in Cancellable(item.Notes, cancellationToken)) yield return new KeyValuePair<string, string>("notes", value);
    }

    private static bool HasBalancedBraces(string value, CancellationToken cancellationToken) {
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value[index] == '\\' && index + 1 < value.Length) { index++; continue; }
            if (value[index] == '{') depth++;
            else if (value[index] == '}' && --depth < 0) return false;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return depth == 0;
    }

    private static bool HasInvalidXmlCharacters(string value, CancellationToken cancellationToken) {
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (char.IsHighSurrogate(value[index])) {
                if (index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) { index++; continue; }
                return true;
            }
            if (char.IsLowSurrogate(value[index])) return true;
            if (!System.Xml.XmlConvert.IsXmlChar(value[index])) return true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return false;
    }

    private static IEnumerable<T> Cancellable<T>(IEnumerable<T> source, CancellationToken cancellationToken) {
        int index = 0;
        foreach (T value in source) {
            if ((index++ & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            yield return value;
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static void Loss(BibliographyConversionReport report, BibliographyItem item, string field, string code, string message, BibliographyConversionAction action) =>
        report.Add(code, BibliographyDiagnosticSeverity.Warning, message, action, item, field);
}
