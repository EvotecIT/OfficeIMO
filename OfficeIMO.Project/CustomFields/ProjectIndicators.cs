using System.Collections.ObjectModel;

namespace OfficeIMO.Project;

/// <summary>Portable graphical indicator symbols for project reports.</summary>
public enum ProjectIndicatorIcon {
    /// <summary>Red status circle.</summary>
    RedCircle,
    /// <summary>Yellow status circle.</summary>
    YellowCircle,
    /// <summary>Green status circle.</summary>
    GreenCircle,
    /// <summary>Blue status circle.</summary>
    BlueCircle,
    /// <summary>Red flag.</summary>
    RedFlag,
    /// <summary>Green flag.</summary>
    GreenFlag
}

/// <summary>An ordered, caller-supplied indicator rule. Native Project graphical-indicator tables are not inferred or rewritten.</summary>
public sealed class ProjectIndicatorRule {
    /// <summary>Creates a predicate using the same bounded expression language as custom fields.</summary>
    public ProjectIndicatorRule(string expression, ProjectIndicatorIcon icon, string label) {
        if (string.IsNullOrWhiteSpace(expression)) throw new ArgumentException("An indicator requires a predicate.", nameof(expression));
        if (!Enum.IsDefined(typeof(ProjectIndicatorIcon), icon)) throw new ArgumentOutOfRangeException(nameof(icon));
        Expression = expression; Icon = icon; Label = label ?? throw new ArgumentNullException(nameof(label));
    }
    /// <summary>Predicate evaluated against the task or resource and its calculated custom fields.</summary>
    public string Expression { get; }
    /// <summary>Symbol selected when the predicate is true.</summary>
    public ProjectIndicatorIcon Icon { get; }
    /// <summary>Accessible text describing the selected status.</summary>
    public string Label { get; }
}

/// <summary>The first matching rule for one entity, or an explicit unmatched result.</summary>
public sealed class ProjectEntityIndicator {
    internal ProjectEntityIndicator(int uid, int? index, ProjectIndicatorRule? rule) { EntityUid = uid; RuleIndex = index; Icon = rule?.Icon; Label = rule?.Label; }
    /// <summary>Task or resource UID within the result's entity namespace.</summary>
    public int EntityUid { get; }
    /// <summary>Zero-based matched rule, or null when no rule matched.</summary>
    public int? RuleIndex { get; }
    /// <summary>Selected graphical symbol, or null.</summary>
    public ProjectIndicatorIcon? Icon { get; }
    /// <summary>Selected accessible label, or null.</summary>
    public string? Label { get; }
}

/// <summary>Immutable indicator evaluation. Errors mark an incomplete result and must be checked before rendering.</summary>
public sealed class ProjectIndicatorResult {
    internal ProjectIndicatorResult(long revision, ProjectCustomFieldEntityKind kind, IEnumerable<ProjectEntityIndicator> values, IEnumerable<ProjectDiagnostic> diagnostics) {
        ModelRevision = revision; EntityKind = kind; Values = new ReadOnlyCollection<ProjectEntityIndicator>(values.ToArray()); Report = new ProjectReport(revision, diagnostics);
    }
    /// <summary>Source document revision.</summary>
    public long ModelRevision { get; }
    /// <summary>Task or resource namespace.</summary>
    public ProjectCustomFieldEntityKind EntityKind { get; }
    /// <summary>Evaluated entities in document order.</summary>
    public IReadOnlyList<ProjectEntityIndicator> Values { get; }
    /// <summary>Expression, input and bounded-evaluation diagnostics.</summary>
    public ProjectReport Report { get; }
}
