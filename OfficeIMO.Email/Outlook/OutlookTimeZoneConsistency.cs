namespace OfficeIMO.Email;

/// <summary>Overall relationship between legacy and definition-based Outlook time-zone values.</summary>
public enum OutlookTimeZoneConsistencyStatus {
    /// <summary>The effective conversion fields match.</summary>
    Consistent,
    /// <summary>One or both representations are absent.</summary>
    Incomplete,
    /// <summary>One or both retained representations could not be decoded.</summary>
    Undecodable,
    /// <summary>The effective conversion fields disagree.</summary>
    Inconsistent
}

/// <summary>One field-level Outlook time-zone consistency issue.</summary>
public sealed class OutlookTimeZoneConsistencyIssue {
    internal OutlookTimeZoneConsistencyIssue(string field, string message) { Field = field; Message = message; }
    /// <summary>Compared field.</summary>
    public string Field { get; }
    /// <summary>Mismatch or missing-state explanation.</summary>
    public string Message { get; }
}

/// <summary>Structured comparison of PidLidTimeZoneStruct and TZDEFINITION.</summary>
public sealed class OutlookTimeZoneConsistencyReport {
    internal OutlookTimeZoneConsistencyReport(OutlookTimeZoneConsistencyStatus status,
        IReadOnlyList<OutlookTimeZoneConsistencyIssue> issues) { Status = status; Issues = issues; }
    /// <summary>Overall comparison status.</summary>
    public OutlookTimeZoneConsistencyStatus Status { get; }
    /// <summary>Field-level evidence.</summary>
    public IReadOnlyList<OutlookTimeZoneConsistencyIssue> Issues { get; }
    /// <summary>Whether both representations decoded and agree.</summary>
    public bool IsConsistent => Status == OutlookTimeZoneConsistencyStatus.Consistent;
}

/// <summary>Consistency checks required by MS-OXOCAL precedence rules.</summary>
public static class OutlookTimeZoneConsistency {
    /// <summary>Compares the legacy structure with the rule effective for <paramref name="localYear"/>.</summary>
    public static OutlookTimeZoneConsistencyReport Compare(OutlookTimeZoneStructure? legacy,
        OutlookTimeZoneDefinition? definition, int localYear) {
        if (localYear < 1 || localYear > 9999) throw new ArgumentOutOfRangeException(nameof(localYear));
        var issues = new List<OutlookTimeZoneConsistencyIssue>();
        if (legacy == null || definition == null) {
            if (legacy == null) issues.Add(new OutlookTimeZoneConsistencyIssue("PidLidTimeZoneStruct",
                "The legacy time-zone structure is absent."));
            if (definition == null) issues.Add(new OutlookTimeZoneConsistencyIssue("TZDEFINITION",
                "The definition-based time-zone property is absent."));
            return new OutlookTimeZoneConsistencyReport(OutlookTimeZoneConsistencyStatus.Incomplete, issues);
        }
        if (!legacy.StateDecoded || !definition.StateDecoded) {
            if (!legacy.StateDecoded) issues.Add(new OutlookTimeZoneConsistencyIssue("PidLidTimeZoneStruct",
                legacy.DecodeError ?? "The legacy time-zone structure did not decode."));
            if (!definition.StateDecoded) issues.Add(new OutlookTimeZoneConsistencyIssue("TZDEFINITION",
                definition.DecodeError ?? "The definition-based time-zone property did not decode."));
            return new OutlookTimeZoneConsistencyReport(OutlookTimeZoneConsistencyStatus.Undecodable, issues);
        }
        OutlookTimeZoneRule effective;
        try {
            effective = definition.GetRule(localYear);
        } catch (InvalidOperationException exception) {
            issues.Add(new OutlookTimeZoneConsistencyIssue("TZDEFINITION", exception.Message));
            return new OutlookTimeZoneConsistencyReport(OutlookTimeZoneConsistencyStatus.Undecodable, issues);
        }
        Compare(issues, "BiasMinutes", legacy.Rule.BiasMinutes, effective.BiasMinutes);
        Compare(issues, "StandardBiasMinutes", legacy.Rule.StandardBiasMinutes,
            effective.StandardBiasMinutes);
        Compare(issues, "DaylightBiasMinutes", legacy.Rule.DaylightBiasMinutes,
            effective.DaylightBiasMinutes);
        if (!legacy.Rule.StandardTransition.Equals(effective.StandardTransition))
            issues.Add(new OutlookTimeZoneConsistencyIssue("StandardTransition",
                "The legacy and definition-based standard transitions differ."));
        if (!legacy.Rule.DaylightTransition.Equals(effective.DaylightTransition))
            issues.Add(new OutlookTimeZoneConsistencyIssue("DaylightTransition",
                "The legacy and definition-based daylight transitions differ."));
        return new OutlookTimeZoneConsistencyReport(issues.Count == 0
            ? OutlookTimeZoneConsistencyStatus.Consistent
            : OutlookTimeZoneConsistencyStatus.Inconsistent, issues);
    }

    private static void Compare(ICollection<OutlookTimeZoneConsistencyIssue> issues, string field,
        int legacy, int definition) {
        if (legacy == definition) return;
        issues.Add(new OutlookTimeZoneConsistencyIssue(field, string.Concat("Legacy value ",
            legacy.ToString(CultureInfo.InvariantCulture), " differs from definition value ",
            definition.ToString(CultureInfo.InvariantCulture), ".")));
    }
}
