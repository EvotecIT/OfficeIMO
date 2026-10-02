namespace OfficeIMO.IWork;

/// <summary>Readable modern cell selectors whose referenced content is not assessed or reconstructed.</summary>
[Flags]
public enum IWorkCellUnsupportedFeatures {
    /// <summary>No selectors in this bounded feature inventory were found. This does not establish complete cell fidelity.</summary>
    None = 0,
    /// <summary>A conditional-style selector is present; rules and effective formatting remain unassessed.</summary>
    ConditionalStyle = 1,
    /// <summary>An applied conditional-rule selector is present; its effective formatting remains unassessed.</summary>
    AppliedConditionalRule = 2,
    /// <summary>A comment selector is unresolved or has unsupported content, including replies.</summary>
    Comment = 4,
    /// <summary>A date/time format selector is present; its display settings remain unassessed.</summary>
    DateFormat = 8,
    /// <summary>A duration format selector is present; its units and display settings remain unassessed.</summary>
    DurationFormat = 16,
    /// <summary>A text format is unresolved or contains settings beyond the qualified default text format.</summary>
    TextFormat = 32,
    /// <summary>A Boolean format is unresolved or contains settings beyond the qualified default Boolean format.</summary>
    BooleanFormat = 64
}
