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
    Comment = 4
}
