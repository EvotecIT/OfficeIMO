namespace OfficeIMO.Project;

internal sealed partial class ProjectCustomFieldCalculator {
    internal ProjectIndicatorResult Indicators(ProjectCustomFieldEntityKind kind, IReadOnlyList<ProjectIndicatorRule> rules, bool includeSummaries) {
        var values = new List<ProjectEntityIndicator>();
        // Populate the same definition resolver and computed-field cache used by normal field calculation.
        var fields = Calculate();
        if (fields.Report.HasErrors) return new ProjectIndicatorResult(_revision, kind, values, fields.Report.Diagnostics);
        IEnumerable<ProjectEntity> entities = kind == ProjectCustomFieldEntityKind.Task ? _document.AllTasks.Cast<ProjectEntity>() : _document.Resources.Cast<ProjectEntity>();
        int operations = 0;
        foreach (var entity in entities) {
            _token.ThrowIfCancellationRequested();
            if (entity is ProjectTask task && (task.IsNull == true || task.IsActive == false || task.IsSummary && !includeSummaries)) continue;
            int? match = null;
            try {
                if (values.Count >= _maxValues) throw new InvalidDataException("Indicator evaluation exceeds MaxValues.");
                for (int index = 0; index < rules.Count; index++) {
                    if (++operations > _maxValues) throw new InvalidDataException("Indicator predicate evaluations exceed MaxValues.");
                    var result = new ProjectFormulaExpression(rules[index].Expression, name => Resolve(entity, name, 0), _maxOperations, _nesting, _culture, _token).Evaluate();
                    if (result.FlagIn(_culture)) { match = index; break; }
                }
                values.Add(new ProjectEntityIndicator(entity.Uid, match, match.HasValue ? rules[match.Value] : null));
            } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException || exception is FormatException ||
                exception is OverflowException || exception is DivideByZeroException || exception is ArgumentException || exception is InvalidOperationException) {
                Error(exception.Message, (kind == ProjectCustomFieldEntityKind.Task ? "/Task" : "/Resource") + "[UID=" + entity.Uid + "]/Indicator");
                if (_diagnostics.Count >= 1000 || operations > _maxValues || values.Count >= _maxValues) break;
            }
        }
        if (_document.Revision != _revision) throw new InvalidOperationException("The document changed during indicator evaluation.");
        return new ProjectIndicatorResult(_revision, kind, values, _diagnostics);
    }
}
