using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Entity scope of a calculated custom field.</summary>
public enum ProjectCustomFieldEntityKind {
    /// <summary>A task or task summary.</summary>
    Task,
    /// <summary>A resource.</summary>
    Resource
}

/// <summary>Bounds for explicit calculation of the supported Project custom-field expression profile.</summary>
public sealed class ProjectCustomFieldCalculationOptions {
    /// <summary>Culture for text conversions, comparisons, and case mapping. Empty means invariant; expression syntax and numeric literals always use invariant XML notation.</summary>
    public string CultureName { get; set; } = "";
    /// <summary>Maximum parser operations for one formula evaluation.</summary>
    public int MaxExpressionOperations { get; set; } = 10000;
    /// <summary>Maximum syntax nesting and dependent-field recursion depth.</summary>
    public int MaxDepth { get; set; } = 64;
    /// <summary>Maximum calculated entity/field pairs across this proposal.</summary>
    public int MaxValues { get; set; } = 100000;
}

/// <summary>A proposed scalar custom-field value in XML-compatible lexical form.</summary>
public sealed class ProjectCalculatedCustomField {
    internal ProjectCalculatedCustomField(ProjectCustomFieldEntityKind kind, int uid, string fieldId, string value) {
        EntityKind = kind; EntityUid = uid; FieldId = fieldId; Value = value;
    }
    /// <summary>Task or resource identity namespace.</summary>
    public ProjectCustomFieldEntityKind EntityKind { get; }
    /// <summary>Stable UID within the selected namespace.</summary>
    public int EntityUid { get; }
    /// <summary>Canonical field ID.</summary>
    public string FieldId { get; }
    /// <summary>Proposed value; dates use local ISO timestamps, flags use 0/1, durations use ISO durations, and numbers use invariant decimal text.</summary>
    public string Value { get; }
}

/// <summary>Immutable custom-field proposal bound to its originating document revision.</summary>
public sealed class ProjectCustomFieldCalculationResult {
    internal readonly ProjectDocument Document;
    internal ProjectCustomFieldCalculationResult(ProjectDocument document, long revision, IEnumerable<ProjectCalculatedCustomField> values, IEnumerable<ProjectDiagnostic> diagnostics) {
        Document = document; ModelRevision = revision;
        Values = new ReadOnlyCollection<ProjectCalculatedCustomField>(values.ToArray());
        Report = new ProjectReport(revision, diagnostics);
    }
    /// <summary>Source model revision.</summary>
    public long ModelRevision { get; }
    /// <summary>Proposed values. Errors prevent applying the entire proposal.</summary>
    public IReadOnlyList<ProjectCalculatedCustomField> Values { get; }
    /// <summary>Unsupported expressions, missing inputs, cycles, and bounded-calculation errors.</summary>
    public ProjectReport Report { get; }
}
