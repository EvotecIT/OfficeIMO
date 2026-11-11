namespace OfficeIMO.Core.Internal;

/// <summary>Collects hidden-data findings and applies the selected removals as the format walks its document.</summary>
internal sealed class OfficeHiddenDataCollector {
    internal const string NotSelectedNote = "Category not selected for removal.";

    private readonly List<OfficeHiddenDataFinding> _findings = new();

    internal OfficeHiddenDataCollector(string format, OfficeHiddenDataOptions? removal) {
        Format = format;
        Removal = removal;
    }

    internal string Format { get; }

    /// <summary>Removal options, or null when only inspecting.</summary>
    internal OfficeHiddenDataOptions? Removal { get; }

    internal bool Removes(OfficeHiddenDataKind kind) => Removal?.Removes(kind) == true;

    internal OfficeTrackedChangesAction TrackedChanges => Removal?.TrackedChanges ?? OfficeTrackedChangesAction.Keep;

    /// <summary>Records a finding. When its category is selected, runs <paramref name="remove"/> and records <paramref name="removedOutcome"/>; a null action retains it with <paramref name="retainedNote"/>.</summary>
    internal void Add(OfficeHiddenDataKind kind, string location, string description, string? value = null,
        Action? remove = null, string? retainedNote = null, OfficeHiddenDataOutcome removedOutcome = OfficeHiddenDataOutcome.Removed) {
        if (Removal == null) {
            Record(kind, location, description, value, OfficeHiddenDataOutcome.Found, null);
        } else if (!Removal.Removes(kind)) {
            Record(kind, location, description, value, OfficeHiddenDataOutcome.Retained, NotSelectedNote);
        } else if (remove == null) {
            Record(kind, location, description, value, OfficeHiddenDataOutcome.Retained, retainedNote ?? "Removal would break the document.");
        } else {
            remove();
            Record(kind, location, description, value, removedOutcome, null);
        }
    }

    /// <summary>Records a finding whose outcome the format decided itself.</summary>
    internal void Record(OfficeHiddenDataKind kind, string location, string description, string? value,
        OfficeHiddenDataOutcome outcome, string? note) =>
        _findings.Add(new OfficeHiddenDataFinding(kind, location, description, value, outcome, note));

    internal OfficeHiddenDataReport Build() => new(Format, _findings.ToList());
}
