using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

public sealed partial class OfflineAddressBookSession {
    /// <summary>
    /// Builds an immutable, bounded index for offline SMTP, proxy, EX/X.500, target, account,
    /// and optional display-name identity resolution.
    /// </summary>
    public OfflineAddressBookIdentityIndex BuildIdentityIndex(
        OfflineAddressBookIdentityIndexOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        OfflineAddressBookIdentityIndexOptions effective = options ??
            new OfflineAddressBookIdentityIndexOptions();
        var matches = new Dictionary<string, List<OfflineAddressBookIdentityCandidate>>(
            StringComparer.OrdinalIgnoreCase);
        var diagnostics = new List<EmailDiagnostic>();
        int originalDiagnosticCount = _diagnostics.Count;
        int entriesScanned = 0;
        int distinctIdentityCount = 0;
        bool identitiesTruncated = false;
        var enumeration = new OfflineAddressBookEnumerationOptions(
            effective.AddressListId, effective.MaxEntries, effective.ContinueOnEntryError);
        foreach (OfflineAddressBookEntry entry in EnumerateEntries(enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            entriesScanned++;
            int retainedForEntry = 0;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            Add(entry.SmtpAddress, "SMTP", OfflineAddressBookIdentityMatchSource.PrimarySmtpAddress);
            Add(entry.X500Address, "EX", OfflineAddressBookIdentityMatchSource.LegacyExchangeDistinguishedName);
            Add(entry.X500Address, "X500", OfflineAddressBookIdentityMatchSource.LegacyExchangeDistinguishedName);
            AddTyped(entry.TargetAddress, OfflineAddressBookIdentityMatchSource.TargetAddress);
            foreach (string proxy in entry.ProxyAddresses) {
                AddTyped(proxy, OfflineAddressBookIdentityMatchSource.ProxyAddress);
            }
            if (effective.IncludeAccountNames) {
                Add(entry.Account, "ACCOUNT", OfflineAddressBookIdentityMatchSource.AccountName);
            }
            if (effective.IncludeDisplayNames) {
                Add(entry.DisplayName, "DISPLAY", OfflineAddressBookIdentityMatchSource.DisplayName);
            }

            void AddTyped(string? sourceValue, OfflineAddressBookIdentityMatchSource source) {
                if (string.IsNullOrWhiteSpace(sourceValue)) return;
                string type;
                string value;
                if (!OfflineAddressBookIdentityIndex.TrySplitTypedValue(sourceValue!, out type, out value)) {
                    value = OfflineAddressBookIdentityIndex.NormalizeValue(sourceValue!);
                    type = value.IndexOf('@') >= 0 ? "SMTP" :
                        value.StartsWith("/o=", StringComparison.OrdinalIgnoreCase) ? "EX" : "OTHER";
                }
                Add(value, type, source);
                if (type == "EX") Add(value, "X500", source);
                else if (type == "X500") Add(value, "EX", source);
            }

            void Add(string? sourceValue, string addressType,
                OfflineAddressBookIdentityMatchSource source) {
                if (string.IsNullOrWhiteSpace(sourceValue)) return;
                string value = OfflineAddressBookIdentityIndex.NormalizeValue(sourceValue!);
                if (value.Length == 0) return;
                string key = OfflineAddressBookIdentityIndex.Key(addressType, value);
                if (!seen.Add(key)) return;
                if (retainedForEntry >= effective.MaxIdentitiesPerEntry) {
                    identitiesTruncated = true;
                    return;
                }
                retainedForEntry++;
                var candidate = new OfflineAddressBookIdentityCandidate(entry, value,
                    OfflineAddressBookIdentityIndex.NormalizeType(addressType), source);
                if (!matches.TryGetValue(key, out List<OfflineAddressBookIdentityCandidate>? candidates)) {
                    candidates = new List<OfflineAddressBookIdentityCandidate>();
                    matches.Add(key, candidates);
                    distinctIdentityCount++;
                }
                candidates.Add(candidate);
            }
        }

        long expectedEntries = effective.AddressListId == null
            ? DeclaredEntryCount
            : AddressLists.Where(list => string.Equals(list.Id, effective.AddressListId,
                StringComparison.Ordinal)).Sum(list => list.DeclaredEntryCount);
        bool entriesComplete = entriesScanned >= expectedEntries && expectedEntries <= effective.MaxEntries;
        if (!entriesComplete) {
            diagnostics.Add(new EmailDiagnostic(
                "OAB_IDENTITY_INDEX_ENTRY_LIMIT",
                string.Concat("The identity index covered ", entriesScanned.ToString(CultureInfo.InvariantCulture),
                    " of ", expectedEntries.ToString(CultureInfo.InvariantCulture), " declared entries."),
                EmailDiagnosticSeverity.Warning,
                SourcePath));
        }
        if (identitiesTruncated) {
            diagnostics.Add(new EmailDiagnostic(
                "OAB_IDENTITY_INDEX_VALUE_LIMIT",
                "At least one entry exceeded MaxIdentitiesPerEntry; its remaining aliases were not indexed.",
                EmailDiagnosticSeverity.Warning,
                SourcePath));
        }
        if (_diagnostics.Count > originalDiagnosticCount) {
            diagnostics.AddRange(_diagnostics.Skip(originalDiagnosticCount));
        }
        bool isComplete = entriesComplete && !identitiesTruncated &&
            !diagnostics.Any(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
        return new OfflineAddressBookIdentityIndex(matches, entriesScanned, distinctIdentityCount,
            isComplete, diagnostics, effective.IncludeAccountNames, effective.IncludeDisplayNames);
    }
}
