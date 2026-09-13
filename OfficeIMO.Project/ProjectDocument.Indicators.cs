namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>
    /// Evaluates ordered caller-supplied graphical rules without mutation. The first true predicate selects the icon and accessible label.
    /// Uses stored built-in values and the custom-field calculator; it neither reads nor writes opaque native graphical-indicator definitions.
    /// MaxValues bounds both returned entities and total predicate evaluations. At most 1000 rules are accepted.
    /// </summary>
    public ProjectIndicatorResult EvaluateIndicators(ProjectCustomFieldEntityKind entityKind, IEnumerable<ProjectIndicatorRule> rules,
        bool includeSummaries = true, ProjectCustomFieldCalculationOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (_batchDepth != 0) throw new InvalidOperationException("Finish the update scope before evaluating indicators.");
        if (!Enum.IsDefined(typeof(ProjectCustomFieldEntityKind), entityKind)) throw new ArgumentOutOfRangeException(nameof(entityKind));
        if (rules == null) throw new ArgumentNullException(nameof(rules));
        var snapshot = new List<ProjectIndicatorRule>();
        foreach (var rule in rules) {
            cancellationToken.ThrowIfCancellationRequested();
            if (rule == null) throw new ArgumentException("Rules cannot contain null.", nameof(rules));
            if (snapshot.Count == 1000) throw new ArgumentException("At most 1000 indicator rules are supported.", nameof(rules));
            snapshot.Add(rule);
        }
        EnsureNotDisposed();
        if (_batchDepth != 0) throw new InvalidOperationException("The rule iterator left an update scope open.");
        options ??= new ProjectCustomFieldCalculationOptions();
        if (options.MaxDepth < 1 || options.MaxDepth > 128 || options.MaxExpressionOperations < 1 || options.MaxExpressionOperations > 1000000 || options.MaxValues < 1)
            throw new ArgumentOutOfRangeException(nameof(options));
        return new ProjectCustomFieldCalculator(this, options, cancellationToken).Indicators(entityKind, snapshot, includeSummaries);
    }
}
