namespace OfficeIMO.Email;

/// <summary>One member of an Outlook personal distribution list.</summary>
public sealed class OutlookDistributionListMember {
    /// <summary>Raw direct or wrapped member EntryID.</summary>
    public byte[]? EntryId { get; set; }

    /// <summary>Raw synchronized One-Off EntryID used to display and address this member.</summary>
    public byte[]? OneOffEntryId { get; set; }

    /// <summary>Decoded member identity when a valid One-Off EntryID is available.</summary>
    public EmailAddress? Address { get; set; }

    /// <summary>Classified direct or wrapped member EntryID family.</summary>
    public OutlookEntryIdKind Kind { get; internal set; }

    /// <summary>Decode warning retained for malformed or unsupported source EntryIDs.</summary>
    public string? DecodeError { get; internal set; }
}

/// <summary>Validation issue for synchronized personal distribution-list member properties.</summary>
public sealed class OutlookDistributionListValidationIssue {
    internal OutlookDistributionListValidationIssue(string code, string message, int? memberIndex = null,
        bool isError = true) {
        Code = code;
        Message = message;
        MemberIndex = memberIndex;
        IsError = isError;
    }

    /// <summary>Stable issue code.</summary>
    public string Code { get; }

    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }

    /// <summary>Zero-based member index, or null for a list-level issue.</summary>
    public int? MemberIndex { get; }

    /// <summary>Whether the issue prevents standards-conforming writing.</summary>
    public bool IsError { get; }
}

/// <summary>Validation report for an Outlook personal distribution list.</summary>
public sealed class OutlookDistributionListValidationReport {
    internal OutlookDistributionListValidationReport(IReadOnlyList<OutlookDistributionListValidationIssue> issues) {
        Issues = issues;
    }

    /// <summary>Whether the list can be written as synchronized Outlook properties.</summary>
    public bool IsValid => !Issues.Any(issue => issue.IsError);

    /// <summary>Validation issues.</summary>
    public IReadOnlyList<OutlookDistributionListValidationIssue> Issues { get; }
}

/// <summary>Typed Outlook personal distribution list with synchronized raw EntryID evidence.</summary>
public sealed class OutlookDistributionList {
    /// <summary>Maximum byte total allowed for either multivalue member property.</summary>
    public const int MaximumMemberPropertyBytes = 15_000;

    private readonly List<OutlookDistributionListMember> _members = new List<OutlookDistributionListMember>();

    /// <summary>Personal distribution-list name.</summary>
    public string? Name { get; set; }

    /// <summary>Source checksum, or null when it was absent. Writers recompute it from <see cref="Members"/>.</summary>
    public int? Checksum { get; set; }

    /// <summary>Members in synchronized property order.</summary>
    public IList<OutlookDistributionListMember> Members => _members;

    /// <summary>Adds a portable one-off member and returns its editable model.</summary>
    public OutlookDistributionListMember Add(EmailAddress address) {
        if (address == null) throw new ArgumentNullException(nameof(address));
        byte[] entryId = OutlookEntryIdCodec.EncodeOneOff(address);
        var member = new OutlookDistributionListMember {
            EntryId = entryId,
            OneOffEntryId = (byte[])entryId.Clone(),
            Address = address,
            Kind = OutlookEntryIdKind.OneOff
        };
        _members.Add(member);
        return member;
    }

    /// <summary>Validates synchronization, size limits, EntryIDs, and an optional retained checksum.</summary>
    public OutlookDistributionListValidationReport Validate() {
        var issues = new List<OutlookDistributionListValidationIssue>();
        var memberEntryIds = new List<byte[]>(_members.Count);
        int memberBytes = 0;
        int oneOffBytes = 0;
        for (int index = 0; index < _members.Count; index++) {
            OutlookDistributionListMember member = _members[index];
            byte[]? entryId = member.EntryId;
            byte[]? oneOff = member.OneOffEntryId;
            if (entryId == null && member.Address != null) entryId = OutlookEntryIdCodec.EncodeOneOff(member.Address);
            if (oneOff == null && member.Address != null) oneOff = OutlookEntryIdCodec.EncodeOneOff(member.Address);
            if (entryId == null || entryId.Length == 0) {
                issues.Add(new OutlookDistributionListValidationIssue(
                    "OUTLOOK_DISTLIST_MEMBER_ENTRYID_REQUIRED",
                    "A distribution-list member requires an EntryID or a portable address.", index));
            } else {
                memberEntryIds.Add(entryId);
                memberBytes = AddBounded(memberBytes, entryId.Length);
            }
            if (oneOff == null || oneOff.Length == 0) {
                issues.Add(new OutlookDistributionListValidationIssue(
                    "OUTLOOK_DISTLIST_ONEOFF_REQUIRED",
                    "A legacy distribution-list member has no synchronized One-Off EntryID or portable address.",
                    index, isError: false));
            } else {
                oneOffBytes = AddBounded(oneOffBytes, oneOff.Length);
                if (!OutlookEntryIdCodec.TryDecodeOneOff(oneOff, out _, out string? error)) {
                    issues.Add(new OutlookDistributionListValidationIssue(
                        "OUTLOOK_DISTLIST_ONEOFF_INVALID",
                        error ?? "The synchronized One-Off EntryID is invalid.", index));
                }
            }
        }
        if (memberBytes >= MaximumMemberPropertyBytes) {
            issues.Add(new OutlookDistributionListValidationIssue(
                "OUTLOOK_DISTLIST_MEMBERS_TOO_LARGE",
                "PidLidDistributionListMembers must be smaller than 15,000 bytes."));
        }
        if (oneOffBytes >= MaximumMemberPropertyBytes) {
            issues.Add(new OutlookDistributionListValidationIssue(
                "OUTLOOK_DISTLIST_ONEOFF_MEMBERS_TOO_LARGE",
                "PidLidDistributionListOneOffMembers must be smaller than 15,000 bytes."));
        }
        if (Checksum.HasValue && memberEntryIds.Count == _members.Count &&
            CalculateChecksum(memberEntryIds) != Checksum.Value) {
            issues.Add(new OutlookDistributionListValidationIssue(
                "OUTLOOK_DISTLIST_CHECKSUM_MISMATCH",
                "The retained distribution-list checksum does not match the ordered member EntryIDs and will be recomputed.",
                isError: false));
        }
        return new OutlookDistributionListValidationReport(issues.AsReadOnly());
    }

    /// <summary>Calculates the MS-OXOCNTC seed-zero IEEE CRC over ordered member EntryID bytes.</summary>
    public static int CalculateChecksum(IEnumerable<byte[]> memberEntryIds) {
        if (memberEntryIds == null) throw new ArgumentNullException(nameof(memberEntryIds));
        uint checksum = 0;
        foreach (byte[] entryId in memberEntryIds) {
            if (entryId == null) throw new ArgumentException("A member EntryID collection cannot contain null.", nameof(memberEntryIds));
            foreach (byte value in entryId) {
                checksum ^= value;
                for (int bit = 0; bit < 8; bit++) {
                    checksum = (checksum & 1) != 0 ? (checksum >> 1) ^ 0xEDB88320u : checksum >> 1;
                }
            }
        }
        return unchecked((int)checksum);
    }

    internal static OutlookDistributionList Project(MapiPropertyBag properties) {
        object[] members = properties.GetValueOrDefault(MapiKnownProperties.PidLid.DistributionListMembers) ?? Array.Empty<object>();
        object[] oneOffMembers = properties.GetValueOrDefault(MapiKnownProperties.PidLid.DistributionListOneOffMembers) ?? Array.Empty<object>();
        int count = Math.Max(members.Length, oneOffMembers.Length);
        var list = new OutlookDistributionList {
            Name = properties.GetValueOrDefault(MapiKnownProperties.PidLid.DistributionListName) ??
                properties.GetValueOrDefault(MapiKnownProperties.PidTag.DisplayName),
            Checksum = properties.GetNullableValue(MapiKnownProperties.PidLid.DistributionListChecksum)
        };
        for (int index = 0; index < count; index++) {
            byte[]? entryId = index < members.Length ? members[index] as byte[] : null;
            byte[]? oneOff = index < oneOffMembers.Length ? oneOffMembers[index] as byte[] : null;
            var member = new OutlookDistributionListMember {
                EntryId = entryId,
                OneOffEntryId = oneOff,
                Kind = entryId == null ? OutlookEntryIdKind.Unknown : OutlookEntryIdCodec.Classify(entryId)
            };
            byte[]? decodable = oneOff ?? (member.Kind == OutlookEntryIdKind.OneOff ? entryId : null);
            if (decodable != null) {
                if (OutlookEntryIdCodec.TryDecodeOneOff(decodable, out EmailAddress? address, out string? error)) {
                    member.Address = address;
                } else {
                    member.DecodeError = error;
                }
            }
            list.Members.Add(member);
        }
        return list;
    }

    internal void WriteTo(MsgPropertyBuilder properties) {
        OutlookDistributionListValidationReport validation = Validate();
        if (!validation.IsValid) {
            throw new InvalidOperationException(string.Join(" ", validation.Issues.Select(issue =>
                string.Concat(issue.Code, ": ", issue.Message))));
        }
        var members = new object[_members.Count];
        bool canWriteOneOff = _members.All(member => member.OneOffEntryId != null || member.Address != null);
        object[]? oneOff = canWriteOneOff ? new object[_members.Count] : null;
        for (int index = 0; index < _members.Count; index++) {
            OutlookDistributionListMember member = _members[index];
            members[index] = member.EntryId ?? OutlookEntryIdCodec.EncodeOneOff(member.Address!);
            if (oneOff != null) oneOff[index] = member.OneOffEntryId ?? OutlookEntryIdCodec.EncodeOneOff(member.Address!);
        }
        properties.Set(MapiKnownProperties.PidLid.DistributionListName, Name);
        properties.Set(MapiKnownProperties.PidTag.DisplayName, Name);
        properties.Set(MapiKnownProperties.PidLid.DistributionListMembers, members);
        properties.Set(MapiKnownProperties.PidLid.DistributionListOneOffMembers, oneOff);
        properties.Set(MapiKnownProperties.PidLid.DistributionListChecksum,
            CalculateChecksum(members.Cast<byte[]>()));
    }

    private static int AddBounded(int current, int value) =>
        current > int.MaxValue - value ? int.MaxValue : current + value;
}
