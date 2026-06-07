namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight metadata for a PDF optional-content group/layer.
/// </summary>
public sealed class PdfOptionalContentGroup {
    internal PdfOptionalContentGroup(
        int? objectNumber,
        string name,
        IReadOnlyList<string> intents,
        bool? isInitiallyVisible,
        bool isLocked,
        bool isInDefaultOrder,
        string? viewState,
        string? printState,
        string? exportState,
        string? usageCreator,
        string? usageSubtype) {
        ObjectNumber = objectNumber;
        Name = name;
        Intents = intents;
        IsInitiallyVisible = isInitiallyVisible;
        IsLocked = isLocked;
        IsInDefaultOrder = isInDefaultOrder;
        ViewState = viewState;
        PrintState = printState;
        ExportState = exportState;
        UsageCreator = usageCreator;
        UsageSubtype = usageSubtype;
    }

    /// <summary>Optional-content group object number when the group is indirect.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Layer display name from the OCG /Name entry.</summary>
    public string Name { get; }

    /// <summary>OCG intent names from /Intent.</summary>
    public IReadOnlyList<string> Intents { get; }

    /// <summary>Initial visibility inferred from the default optional-content configuration.</summary>
    public bool? IsInitiallyVisible { get; }

    /// <summary>True when the group appears in the default configuration /Locked list.</summary>
    public bool IsLocked { get; }

    /// <summary>True when the group appears in the default configuration /Order list.</summary>
    public bool IsInDefaultOrder { get; }

    /// <summary>Usage /View /ViewState value, when present.</summary>
    public string? ViewState { get; }

    /// <summary>Usage /Print /PrintState value, when present.</summary>
    public string? PrintState { get; }

    /// <summary>Usage /Export /ExportState value, when present.</summary>
    public string? ExportState { get; }

    /// <summary>Usage /CreatorInfo /Creator value, when present.</summary>
    public string? UsageCreator { get; }

    /// <summary>Usage /CreatorInfo /Subtype value, when present.</summary>
    public string? UsageSubtype { get; }
}
