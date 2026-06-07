namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight metadata from a signature field /SV seed value dictionary.
/// </summary>
public sealed class PdfSignatureSeedValueInfo {
    internal PdfSignatureSeedValueInfo(
        string? filter,
        IReadOnlyList<string> subFilters,
        IReadOnlyList<string> digestMethods,
        IReadOnlyList<string> reasons,
        int? flags,
        bool? addRevInfo,
        int? mdpPermissionLevel) {
        Filter = filter;
        SubFilters = subFilters;
        DigestMethods = digestMethods;
        Reasons = reasons;
        Flags = flags;
        AddRevInfo = addRevInfo;
        MDPPermissionLevel = mdpPermissionLevel;
    }

    /// <summary>Preferred or required signature handler /Filter name, when readable.</summary>
    public string? Filter { get; }

    /// <summary>Allowed or required /SubFilter names, when readable.</summary>
    public IReadOnlyList<string> SubFilters { get; }

    /// <summary>Allowed or required /DigestMethod names, when readable.</summary>
    public IReadOnlyList<string> DigestMethods { get; }

    /// <summary>Allowed or suggested signing reasons, when readable.</summary>
    public IReadOnlyList<string> Reasons { get; }

    /// <summary>Raw seed value /Ff flags, when readable.</summary>
    public int? Flags { get; }

    /// <summary>Seed value /AddRevInfo flag, when readable.</summary>
    public bool? AddRevInfo { get; }

    /// <summary>Seed value /MDP /P permission level, when readable.</summary>
    public int? MDPPermissionLevel { get; }
}
