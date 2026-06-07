namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight metadata for catalog /OCProperties optional-content/layer state.
/// </summary>
public sealed class PdfOptionalContentProperties {
    internal PdfOptionalContentProperties(
        IReadOnlyList<PdfOptionalContentGroup> groups,
        string? defaultConfigurationName,
        string? defaultConfigurationCreator,
        string? baseState,
        IReadOnlyList<int> onGroupObjectNumbers,
        IReadOnlyList<int> offGroupObjectNumbers,
        IReadOnlyList<int> lockedGroupObjectNumbers,
        IReadOnlyList<int> orderGroupObjectNumbers) {
        Groups = groups;
        DefaultConfigurationName = defaultConfigurationName;
        DefaultConfigurationCreator = defaultConfigurationCreator;
        BaseState = baseState;
        OnGroupObjectNumbers = onGroupObjectNumbers;
        OffGroupObjectNumbers = offGroupObjectNumbers;
        LockedGroupObjectNumbers = lockedGroupObjectNumbers;
        OrderGroupObjectNumbers = orderGroupObjectNumbers;
    }

    /// <summary>Optional-content groups listed in catalog /OCProperties /OCGs.</summary>
    public IReadOnlyList<PdfOptionalContentGroup> Groups { get; }

    /// <summary>Number of optional-content groups discovered.</summary>
    public int GroupCount => Groups.Count;

    /// <summary>True when at least one optional-content group was discovered.</summary>
    public bool HasGroups => GroupCount > 0;

    /// <summary>Default optional-content configuration /Name, when present.</summary>
    public string? DefaultConfigurationName { get; }

    /// <summary>Default optional-content configuration /Creator, when present.</summary>
    public string? DefaultConfigurationCreator { get; }

    /// <summary>Default optional-content configuration /BaseState, when present.</summary>
    public string? BaseState { get; }

    /// <summary>Group object numbers listed in the default configuration /ON array.</summary>
    public IReadOnlyList<int> OnGroupObjectNumbers { get; }

    /// <summary>Group object numbers listed in the default configuration /OFF array.</summary>
    public IReadOnlyList<int> OffGroupObjectNumbers { get; }

    /// <summary>Group object numbers listed in the default configuration /Locked array.</summary>
    public IReadOnlyList<int> LockedGroupObjectNumbers { get; }

    /// <summary>Group object numbers found in the default configuration /Order array.</summary>
    public IReadOnlyList<int> OrderGroupObjectNumbers { get; }
}
