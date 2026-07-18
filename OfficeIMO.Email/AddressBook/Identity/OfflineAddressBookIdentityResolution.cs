using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

/// <summary>Outcome of an offline directory identity lookup.</summary>
public enum OfflineAddressBookIdentityResolutionStatus {
    /// <summary>Exactly one directory entry matched.</summary>
    Resolved,
    /// <summary>More than one directory entry matched and no winner was guessed.</summary>
    Ambiguous,
    /// <summary>No entry matched a complete index.</summary>
    NotFound,
    /// <summary>No entry matched, but the configured index bounds or source errors prevented a complete answer.</summary>
    Incomplete
}

/// <summary>Directory field that supplied an identity match.</summary>
public enum OfflineAddressBookIdentityMatchSource {
    /// <summary>Primary SMTP address.</summary>
    PrimarySmtpAddress,
    /// <summary>Primary legacy Exchange/X.500 distinguished name.</summary>
    LegacyExchangeDistinguishedName,
    /// <summary>Target address.</summary>
    TargetAddress,
    /// <summary>Proxy-address alias.</summary>
    ProxyAddress,
    /// <summary>Directory account name.</summary>
    AccountName,
    /// <summary>Display-name heuristic.</summary>
    DisplayName
}

/// <summary>Per-lookup policy for non-address fallbacks and returned ambiguity evidence.</summary>
public sealed class OfflineAddressBookIdentityResolutionOptions {
    /// <summary>Creates resolution options.</summary>
    public OfflineAddressBookIdentityResolutionOptions(
        bool allowAccountNameMatch = true,
        bool allowDisplayNameMatch = false,
        int maxCandidates = 32) {
        if (maxCandidates <= 0) throw new ArgumentOutOfRangeException(nameof(maxCandidates));
        AllowAccountNameMatch = allowAccountNameMatch;
        AllowDisplayNameMatch = allowDisplayNameMatch;
        MaxCandidates = maxCandidates;
    }

    /// <summary>Whether an untyped non-address value may match an exact directory account name.</summary>
    public bool AllowAccountNameMatch { get; }

    /// <summary>Whether an untyped non-address value may match an exact display name.</summary>
    public bool AllowDisplayNameMatch { get; }

    /// <summary>Maximum candidates retained when the result is ambiguous.</summary>
    public int MaxCandidates { get; }
}

/// <summary>One exact or heuristic directory match, with enough data to resolve EX/X.500 identities to SMTP.</summary>
public sealed class OfflineAddressBookIdentityCandidate {
    internal OfflineAddressBookIdentityCandidate(OfflineAddressBookEntry entry,
        string matchedIdentity, string addressType, OfflineAddressBookIdentityMatchSource source) {
        Reference = entry.Reference;
        DisplayName = entry.DisplayName;
        PrimarySmtpAddress = entry.SmtpAddress;
        LegacyExchangeDistinguishedName = entry.X500Address;
        AccountName = entry.Account;
        ObjectType = entry.ObjectType;
        IsDistributionList = entry.IsDistributionList;
        MatchedIdentity = matchedIdentity;
        AddressType = addressType;
        MatchSource = source;
    }

    /// <summary>Stable reference that can be read through the originating session while it remains open.</summary>
    public OfflineAddressBookEntryReference Reference { get; }
    /// <summary>Directory display name.</summary>
    public string? DisplayName { get; }
    /// <summary>Primary SMTP address, when published by the OAB.</summary>
    public string? PrimarySmtpAddress { get; }
    /// <summary>Primary legacy Exchange/X.500 distinguished name, when published by the OAB.</summary>
    public string? LegacyExchangeDistinguishedName { get; }
    /// <summary>Directory account name, when published by the OAB.</summary>
    public string? AccountName { get; }
    /// <summary>Projected directory object type.</summary>
    public OfflineAddressBookObjectType ObjectType { get; }
    /// <summary>Whether the entry represents a distribution list.</summary>
    public bool IsDistributionList { get; }
    /// <summary>Exact source value that matched.</summary>
    public string MatchedIdentity { get; }
    /// <summary>Normalized address type used for the match, such as SMTP, EX, X500, SIP, ACCOUNT, or DISPLAY.</summary>
    public string AddressType { get; }
    /// <summary>Directory field that supplied the match.</summary>
    public OfflineAddressBookIdentityMatchSource MatchSource { get; }
    /// <summary>True for address-valued directory fields; false for account and display-name fallbacks.</summary>
    public bool IsAuthoritativeAddress => MatchSource != OfflineAddressBookIdentityMatchSource.AccountName &&
        MatchSource != OfflineAddressBookIdentityMatchSource.DisplayName;

    /// <summary>Returns the best portable address, preferring primary SMTP over the legacy Exchange identity.</summary>
    public EmailAddress ToEmailAddress() {
        if (!string.IsNullOrWhiteSpace(PrimarySmtpAddress)) {
            return new EmailAddress(PrimarySmtpAddress, DisplayName) { AddressType = "SMTP" };
        }
        return new EmailAddress(LegacyExchangeDistinguishedName ?? MatchedIdentity, DisplayName) {
            AddressType = string.Equals(AddressType, "X500", StringComparison.Ordinal) ? "EX" : AddressType
        };
    }
}

/// <summary>Duplicate-aware identity-resolution result.</summary>
public sealed class OfflineAddressBookIdentityResolution {
    internal OfflineAddressBookIdentityResolution(OfflineAddressBookIdentityResolutionStatus status,
        IReadOnlyList<OfflineAddressBookIdentityCandidate> candidates, bool candidatesTruncated,
        bool indexIsComplete) {
        Status = status;
        Candidates = candidates;
        CandidatesTruncated = candidatesTruncated;
        IndexIsComplete = indexIsComplete;
    }

    /// <summary>Resolution outcome.</summary>
    public OfflineAddressBookIdentityResolutionStatus Status { get; }
    /// <summary>Unique matching directory entries in deterministic order.</summary>
    public IReadOnlyList<OfflineAddressBookIdentityCandidate> Candidates { get; }
    /// <summary>Whether additional ambiguous candidates were omitted by the per-query bound.</summary>
    public bool CandidatesTruncated { get; }
    /// <summary>Whether the index covered all selected source records and their configured identity values.</summary>
    public bool IndexIsComplete { get; }
    /// <summary>The sole candidate for a resolved result; otherwise null.</summary>
    public OfflineAddressBookIdentityCandidate? Candidate =>
        Status == OfflineAddressBookIdentityResolutionStatus.Resolved && Candidates.Count == 1
            ? Candidates[0]
            : null;
}
