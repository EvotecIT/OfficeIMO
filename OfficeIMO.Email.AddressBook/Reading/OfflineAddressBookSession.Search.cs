using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

public sealed partial class OfflineAddressBookSession {
    /// <summary>
    /// Searches semantic address-entry fields with explicit record, character, result, progress, cancellation, and
    /// exact-position resume bounds.
    /// </summary>
    public OfflineAddressBookSearchReport Search(
        OfflineAddressBookSearchQuery query,
        IProgress<OfflineAddressBookSearchProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        if (query == null) throw new ArgumentNullException(nameof(query));
        ThrowIfDisposed();

        IReadOnlyList<OabAddressListSource> selectedSources = SelectSearchSources(query);
        SearchPosition position = GetSearchStart(selectedSources, query.ResumeFrom);
        var results = new List<OfflineAddressBookSearchResult>();
        var diagnostics = new List<EmailDiagnostic>();
        int scanned = 0;
        int skipped = 0;
        bool stoppedAtEntryLimit = false;
        bool stoppedAtResultLimit = false;
        OfflineAddressBookSearchCheckpoint? nextCheckpoint = null;

        for (int selectedIndex = position.SelectedSourceIndex;
            selectedIndex < selectedSources.Count; selectedIndex++) {
            OabAddressListSource selected = selectedSources[selectedIndex];
            long entryIndex = selectedIndex == position.SelectedSourceIndex ? position.EntryIndex : 0;
            long recordOffset = selectedIndex == position.SelectedSourceIndex
                ? position.RecordOffset
                : selected.Info.EntriesOffset;
            using (OabStreamLease lease = selected.Source.OpenRead()) {
                Stream stream = lease.Stream;
                OabBinary.Seek(selected.Source, stream, recordOffset, selected.Info.Id);
                while (entryIndex < selected.Info.DeclaredEntryCount) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (scanned >= query.MaxEntriesScanned) {
                        stoppedAtEntryLimit = true;
                        nextCheckpoint = CreateCheckpoint(selected, entryIndex,
                            stream.Position - selected.Source.BaseOffset);
                        break;
                    }

                    long currentOffset = stream.Position - selected.Source.BaseOffset;
                    string location = BuildEntryLocation(selected.Info, entryIndex);
                    OabRecordEnvelope envelope;
                    try {
                        envelope = OabV4RecordReader.ReadEnvelope(selected.Source, stream, _options, location);
                    } catch (Exception exception) when (
                        query.ContinueOnEntryError && IsRecoverableReadException(exception)) {
                        skipped++;
                        scanned++;
                        diagnostics.Add(SearchDiagnostic("OAB_SEARCH_FRAMING_STOPPED", exception, location));
                        entryIndex = selected.Info.DeclaredEntryCount;
                        ReportSearchProgressIfDue(query, progress, scanned, results.Count, skipped);
                        break;
                    }

                    var reference = new OfflineAddressBookEntryReference(
                        selected.Info.Id, selected.Info.Index, entryIndex, currentOffset, envelope.Size,
                        selected.SnapshotId);
                    entryIndex++;
                    scanned++;
                    try {
                        OabParsedRecord record = OabV4RecordReader.Parse(
                            envelope, selected.Info.EntryPropertyDefinitions, _options, reference.Id);
                        var entry = new OfflineAddressBookEntry(
                            reference, selected.Info, record.Properties, record.Diagnostics);
                        if ((!query.ObjectType.HasValue || entry.ObjectType == query.ObjectType.Value) &&
                            TryMatch(entry, query, out OfflineAddressBookSearchFields matchedFields,
                                out string? snippet)) {
                            results.Add(new OfflineAddressBookSearchResult(
                                new OfflineAddressBookEntrySummary(entry), matchedFields, snippet));
                            if (results.Count >= query.MaxResults &&
                                HasMore(selectedSources, selectedIndex, entryIndex)) {
                                stoppedAtResultLimit = true;
                                nextCheckpoint = CreateNextCheckpoint(
                                    selectedSources, selectedIndex, selected, entryIndex,
                                    stream.Position - selected.Source.BaseOffset);
                                break;
                            }
                        }
                    } catch (Exception exception) when (
                        query.ContinueOnEntryError && IsRecoverableReadException(exception)) {
                        skipped++;
                        diagnostics.Add(SearchDiagnostic("OAB_SEARCH_ENTRY_SKIPPED", exception, reference.Id));
                    }
                    ReportSearchProgressIfDue(query, progress, scanned, results.Count, skipped);
                }
            }
            if (nextCheckpoint != null) break;
        }

        ReportSearchProgress(progress, scanned, results.Count, skipped);
        return new OfflineAddressBookSearchReport(
            results, diagnostics, scanned, skipped,
            stoppedAtEntryLimit, stoppedAtResultLimit, nextCheckpoint);
    }

    private IReadOnlyList<OabAddressListSource> SelectSearchSources(OfflineAddressBookSearchQuery query) {
        IReadOnlyList<OabAddressListSource> sources = query.AddressListId == null
            ? _sources
            : SelectSources(query.AddressListId).ToArray();
        if (query.ResumeFrom != null && query.AddressListId != null &&
            !string.Equals(query.AddressListId, query.ResumeFrom.AddressListId, StringComparison.Ordinal)) {
            throw new ArgumentException("The search checkpoint is outside the selected address list.", nameof(query));
        }
        return sources;
    }

    private static SearchPosition GetSearchStart(IReadOnlyList<OabAddressListSource> selectedSources,
        OfflineAddressBookSearchCheckpoint? checkpoint) {
        if (checkpoint == null) {
            OabAddressListSource first = selectedSources[0];
            return new SearchPosition(0, 0, first.Info.EntriesOffset);
        }
        for (int index = 0; index < selectedSources.Count; index++) {
            OabAddressListSource source = selectedSources[index];
            if (source.SnapshotId != checkpoint.SnapshotId ||
                source.Info.Index != checkpoint.AddressListIndex ||
                !string.Equals(source.Info.Id, checkpoint.AddressListId, StringComparison.Ordinal)) continue;
            if (checkpoint.EntryIndex < 0 || checkpoint.EntryIndex >= source.Info.DeclaredEntryCount ||
                checkpoint.RecordOffset < source.Info.EntriesOffset ||
                checkpoint.RecordOffset > source.Source.Length - 4) {
                throw new ArgumentException("The search checkpoint is invalid for this session snapshot.", nameof(checkpoint));
            }
            return new SearchPosition(index, checkpoint.EntryIndex, checkpoint.RecordOffset);
        }
        throw new ArgumentException("The search checkpoint is outside the selected session scope.", nameof(checkpoint));
    }

    private static bool HasMore(IReadOnlyList<OabAddressListSource> selectedSources,
        int selectedIndex, long nextEntryIndex) {
        if (nextEntryIndex < selectedSources[selectedIndex].Info.DeclaredEntryCount) return true;
        for (int index = selectedIndex + 1; index < selectedSources.Count; index++) {
            if (selectedSources[index].Info.DeclaredEntryCount > 0) return true;
        }
        return false;
    }

    private static OfflineAddressBookSearchCheckpoint CreateNextCheckpoint(
        IReadOnlyList<OabAddressListSource> selectedSources, int selectedIndex,
        OabAddressListSource selected, long nextEntryIndex, long nextOffset) {
        if (nextEntryIndex < selected.Info.DeclaredEntryCount) {
            return CreateCheckpoint(selected, nextEntryIndex, nextOffset);
        }
        for (int index = selectedIndex + 1; index < selectedSources.Count; index++) {
            OabAddressListSource next = selectedSources[index];
            if (next.Info.DeclaredEntryCount > 0) return CreateCheckpoint(next, 0, next.Info.EntriesOffset);
        }
        throw new InvalidOperationException("A search checkpoint was requested after the selected scope was exhausted.");
    }

    private static OfflineAddressBookSearchCheckpoint CreateCheckpoint(
        OabAddressListSource source, long entryIndex, long recordOffset) =>
        new OfflineAddressBookSearchCheckpoint(
            source.Info.Id, source.Info.Index, entryIndex, recordOffset, source.SnapshotId);

    private static EmailDiagnostic SearchDiagnostic(string code, Exception exception, string location) =>
        new EmailDiagnostic(code, exception.Message,
            exception is OfflineAddressBookLimitExceededException
                ? EmailDiagnosticSeverity.Warning
                : EmailDiagnosticSeverity.Error,
            location);

    private static bool TryMatch(OfflineAddressBookEntry entry,
        OfflineAddressBookSearchQuery query,
        out OfflineAddressBookSearchFields matchedFields,
        out string? snippet) {
        var fields = new List<SearchFieldText>();
        int remaining = query.MaxSearchableCharactersPerEntry;
        if ((query.Fields & OfflineAddressBookSearchFields.Names) != 0) {
            AddSearchField(fields, OfflineAddressBookSearchFields.Names, entry.DisplayName, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Names, entry.GivenName, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Names, entry.Surname, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Names, entry.Initials, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Names, entry.Account, ref remaining);
        }
        if ((query.Fields & OfflineAddressBookSearchFields.Addresses) != 0) {
            AddSearchField(fields, OfflineAddressBookSearchFields.Addresses, entry.SmtpAddress, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Addresses, entry.X500Address, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Addresses, entry.TargetAddress, ref remaining);
            AddSearchFields(fields, OfflineAddressBookSearchFields.Addresses, entry.ProxyAddresses, ref remaining);
        }
        if ((query.Fields & OfflineAddressBookSearchFields.Organization) != 0) {
            AddSearchField(fields, OfflineAddressBookSearchFields.Organization, entry.CompanyName, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Organization, entry.Department, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Organization, entry.JobTitle, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Organization, entry.OfficeLocation, ref remaining);
        }
        if ((query.Fields & OfflineAddressBookSearchFields.Phones) != 0) {
            AddSearchField(fields, OfflineAddressBookSearchFields.Phones, entry.BusinessTelephone, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Phones, entry.HomeTelephone, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Phones, entry.MobileTelephone, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Phones, entry.PrimaryFax, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Phones, entry.AssistantTelephone, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.Phones, entry.PagerTelephone, ref remaining);
            AddSearchFields(fields, OfflineAddressBookSearchFields.Phones, entry.BusinessTelephone2, ref remaining);
            AddSearchFields(fields, OfflineAddressBookSearchFields.Phones, entry.HomeTelephone2, ref remaining);
        }
        if ((query.Fields & OfflineAddressBookSearchFields.PostalAddress) != 0) {
            AddSearchField(fields, OfflineAddressBookSearchFields.PostalAddress, entry.StreetAddress, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.PostalAddress, entry.Locality, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.PostalAddress, entry.StateOrProvince, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.PostalAddress, entry.PostalCode, ref remaining);
            AddSearchField(fields, OfflineAddressBookSearchFields.PostalAddress, entry.Country, ref remaining);
        }
        if ((query.Fields & OfflineAddressBookSearchFields.Comment) != 0) {
            AddSearchField(fields, OfflineAddressBookSearchFields.Comment, entry.Comment, ref remaining);
        }
        if ((query.Fields & OfflineAddressBookSearchFields.Membership) != 0) {
            AddSearchFields(fields, OfflineAddressBookSearchFields.Membership,
                entry.MemberDistinguishedNames, ref remaining);
            AddSearchFields(fields, OfflineAddressBookSearchFields.Membership,
                entry.MemberOfDistinguishedNames, ref remaining);
        }

        var matchedTerms = new bool[query.Terms.Count];
        matchedFields = OfflineAddressBookSearchFields.None;
        string? snippetSource = null;
        foreach (SearchFieldText field in fields) {
            bool fieldMatched = false;
            for (int termIndex = 0; termIndex < query.Terms.Count; termIndex++) {
                if (field.Text.IndexOf(query.Terms[termIndex], StringComparison.OrdinalIgnoreCase) < 0) continue;
                matchedTerms[termIndex] = true;
                fieldMatched = true;
            }
            if (!fieldMatched) continue;
            matchedFields |= field.Field;
            if (snippetSource == null) snippetSource = field.Text;
        }
        bool match = query.MatchMode == OfflineAddressBookSearchMatchMode.AnyTerm
            ? matchedTerms.Any(value => value)
            : matchedTerms.All(value => value);
        snippet = match && snippetSource != null
            ? OabSearchText.CreateSnippet(snippetSource, query.Terms, query.SnippetCharacters)
            : null;
        return match;
    }

    private static void AddSearchFields(List<SearchFieldText> fields,
        OfflineAddressBookSearchFields field, IReadOnlyList<string> values, ref int remaining) {
        foreach (string value in values) AddSearchField(fields, field, value, ref remaining);
    }

    private static void AddSearchField(List<SearchFieldText> fields,
        OfflineAddressBookSearchFields field, string? value, ref int remaining) {
        if (remaining <= 0 || string.IsNullOrEmpty(value)) return;
        string normalized = OabSearchText.Normalize(value, remaining);
        if (normalized.Length == 0) return;
        fields.Add(new SearchFieldText(field, normalized));
        remaining -= normalized.Length;
    }

    private static void ReportSearchProgressIfDue(OfflineAddressBookSearchQuery query,
        IProgress<OfflineAddressBookSearchProgress>? progress,
        int scanned, int matches, int skipped) {
        if (scanned % query.ProgressInterval == 0) {
            ReportSearchProgress(progress, scanned, matches, skipped);
        }
    }

    private static void ReportSearchProgress(IProgress<OfflineAddressBookSearchProgress>? progress,
        int scanned, int matches, int skipped) =>
        progress?.Report(new OfflineAddressBookSearchProgress(scanned, matches, skipped));

    private readonly struct SearchPosition {
        internal SearchPosition(int selectedSourceIndex, long entryIndex, long recordOffset) {
            SelectedSourceIndex = selectedSourceIndex;
            EntryIndex = entryIndex;
            RecordOffset = recordOffset;
        }
        internal int SelectedSourceIndex { get; }
        internal long EntryIndex { get; }
        internal long RecordOffset { get; }
    }

    private readonly struct SearchFieldText {
        internal SearchFieldText(OfflineAddressBookSearchFields field, string text) {
            Field = field;
            Text = text;
        }
        internal OfflineAddressBookSearchFields Field { get; }
        internal string Text { get; }
    }
}
