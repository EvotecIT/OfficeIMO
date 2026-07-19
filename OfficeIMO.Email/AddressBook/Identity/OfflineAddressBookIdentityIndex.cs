using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

/// <summary>
/// Immutable, duplicate-aware offline identity index for resolving SMTP, proxy, EX/X.500, target,
/// account, and optionally display-name identities without Outlook, Exchange, or network access.
/// </summary>
public sealed class OfflineAddressBookIdentityIndex {
    private readonly Dictionary<string, List<OfflineAddressBookIdentityCandidate>> _matches;

    internal OfflineAddressBookIdentityIndex(
        Dictionary<string, List<OfflineAddressBookIdentityCandidate>> matches,
        int entriesScanned, int identityCount, bool isComplete,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        bool includesAccountNames, bool includesDisplayNames) {
        _matches = matches;
        EntriesScanned = entriesScanned;
        IdentityCount = identityCount;
        IsComplete = isComplete;
        Diagnostics = diagnostics;
        IncludesAccountNames = includesAccountNames;
        IncludesDisplayNames = includesDisplayNames;
    }

    /// <summary>Successfully decoded entries scanned while building the index.</summary>
    public int EntriesScanned { get; }
    /// <summary>Distinct typed identity keys retained by the index.</summary>
    public int IdentityCount { get; }
    /// <summary>Whether every selected source entry and configured identity was indexed.</summary>
    public bool IsComplete { get; }
    /// <summary>Build diagnostics, including truncation and recoverable source failures.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>Whether account-name aliases were included.</summary>
    public bool IncludesAccountNames { get; }
    /// <summary>Whether display-name heuristics were included.</summary>
    public bool IncludesDisplayNames { get; }

    /// <summary>Resolves a raw identity and optional address type without guessing between duplicate entries.</summary>
    public OfflineAddressBookIdentityResolution Resolve(string identity,
        string? addressType = null,
        OfflineAddressBookIdentityResolutionOptions? options = null) {
        if (identity == null) throw new ArgumentNullException(nameof(identity));
        OfflineAddressBookIdentityResolutionOptions effective = options ??
            new OfflineAddressBookIdentityResolutionOptions();
        IdentityLookup lookup = IdentityLookup.Create(identity, addressType, effective);
        var byReference = new Dictionary<string, OfflineAddressBookIdentityCandidate>(StringComparer.Ordinal);
        foreach (string key in lookup.Keys) {
            if (!_matches.TryGetValue(key, out List<OfflineAddressBookIdentityCandidate>? found)) continue;
            foreach (OfflineAddressBookIdentityCandidate candidate in found) {
                if (byReference.TryGetValue(candidate.Reference.Id,
                    out OfflineAddressBookIdentityCandidate? current) &&
                    Rank(current.MatchSource) <= Rank(candidate.MatchSource)) continue;
                byReference[candidate.Reference.Id] = candidate;
            }
        }

        OfflineAddressBookIdentityCandidate[] all = byReference.Values
            .OrderBy(candidate => candidate.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(candidate => candidate.PrimarySmtpAddress, StringComparer.OrdinalIgnoreCase)
            .ThenBy(candidate => candidate.Reference.Id, StringComparer.Ordinal)
            .ToArray();
        bool truncated = all.Length > effective.MaxCandidates;
        IReadOnlyList<OfflineAddressBookIdentityCandidate> candidates = truncated
            ? all.Take(effective.MaxCandidates).ToArray()
            : all;
        OfflineAddressBookIdentityResolutionStatus status = all.Length switch {
            0 => IsComplete
                ? OfflineAddressBookIdentityResolutionStatus.NotFound
                : OfflineAddressBookIdentityResolutionStatus.Incomplete,
            1 => OfflineAddressBookIdentityResolutionStatus.Resolved,
            _ => OfflineAddressBookIdentityResolutionStatus.Ambiguous
        };
        return new OfflineAddressBookIdentityResolution(status, candidates, truncated, IsComplete);
    }

    /// <summary>Resolves an <see cref="EmailAddress"/>, honoring its Outlook address type when present.</summary>
    public OfflineAddressBookIdentityResolution Resolve(EmailAddress address,
        OfflineAddressBookIdentityResolutionOptions? options = null) {
        if (address == null) throw new ArgumentNullException(nameof(address));
        string identity = !string.IsNullOrWhiteSpace(address.Address) ? address.Address! :
            !string.IsNullOrWhiteSpace(address.RawValue) ? address.RawValue! :
            address.DisplayName ?? throw new ArgumentException("The address has no identity value.", nameof(address));
        return Resolve(identity, address.AddressType, options);
    }

    internal static int Rank(OfflineAddressBookIdentityMatchSource source) => source switch {
        OfflineAddressBookIdentityMatchSource.PrimarySmtpAddress => 0,
        OfflineAddressBookIdentityMatchSource.LegacyExchangeDistinguishedName => 1,
        OfflineAddressBookIdentityMatchSource.ProxyAddress => 2,
        OfflineAddressBookIdentityMatchSource.TargetAddress => 3,
        OfflineAddressBookIdentityMatchSource.AccountName => 4,
        _ => 5
    };

    internal static string Key(string addressType, string value) => string.Concat(
        NormalizeType(addressType), "\u001F", NormalizeValue(value));

    internal static string NormalizeType(string value) {
        string type = value.Trim().TrimEnd(':').ToUpperInvariant();
        return type == "X.500" ? "X500" : type;
    }

    internal static string NormalizeValue(string value) => value.Trim().Trim('<', '>');

    internal static bool TrySplitTypedValue(string value, out string addressType, out string identity) {
        string trimmed = value.Trim();
        int separator = trimmed.IndexOf(':');
        if (separator > 0 && separator <= 16 && trimmed.IndexOf('@', 0, separator) < 0) {
            addressType = NormalizeType(trimmed.Substring(0, separator));
            identity = NormalizeValue(trimmed.Substring(separator + 1));
            return identity.Length > 0;
        }
        addressType = string.Empty;
        identity = NormalizeValue(trimmed);
        return false;
    }

    private static string InferType(string identity) {
        if (identity.IndexOf('@') >= 0) return "SMTP";
        if (identity.StartsWith("/o=", StringComparison.OrdinalIgnoreCase)) return "EX";
        return string.Empty;
    }

    private readonly struct IdentityLookup {
        private IdentityLookup(IReadOnlyList<string> keys) { Keys = keys; }
        internal IReadOnlyList<string> Keys { get; }

        internal static IdentityLookup Create(string identity, string? addressType,
            OfflineAddressBookIdentityResolutionOptions options) {
            string value = NormalizeValue(identity);
            if (value.Length == 0) throw new ArgumentException("Identity cannot be empty.", nameof(identity));
            string? explicitType = string.IsNullOrWhiteSpace(addressType) ? null : NormalizeType(addressType!);
            if (TrySplitTypedValue(value, out string prefixedType, out string prefixedValue)) {
                explicitType = prefixedType;
                value = prefixedValue;
            }
            var keys = new List<string>();
            string inferred = explicitType ?? InferType(value);
            if (inferred.Length > 0) AddType(keys, inferred, value);
            if (inferred.Length == 0 && options.AllowAccountNameMatch) keys.Add(Key("ACCOUNT", value));
            if (inferred.Length == 0 && options.AllowDisplayNameMatch) keys.Add(Key("DISPLAY", value));
            return new IdentityLookup(keys);
        }

        private static void AddType(ICollection<string> keys, string type, string value) {
            keys.Add(Key(type, value));
            if (type == "EX") keys.Add(Key("X500", value));
            else if (type == "X500") keys.Add(Key("EX", value));
        }
    }
}
