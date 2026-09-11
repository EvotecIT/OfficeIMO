namespace OfficeIMO.Project;

internal sealed partial class ProjectCustomFieldCalculator {
    private ProjectFormulaValue Rollup(ProjectTask task, ProjectCustomFieldIdentity identity, ProjectCustomFieldDefinition definition, int depth) {
        int operation = definition.RollupType ?? throw new InvalidDataException("Summary rollup requires an explicit operator.");
        if (operation < 0 || operation > 7) throw new NotSupportedException("The summary rollup operator is outside the supported profile.");
        if (identity.Kind == "Text") throw new NotSupportedException("Text fields support summary formulas, not rollup operators, in this profile.");
        if (identity.Kind == "Flag" && operation != 0 && operation != 1) throw new NotSupportedException("Flag fields support only OR and AND rollups.");
        if (identity.Kind == "Date" && operation != 0 && operation != 1) throw new NotSupportedException("Date fields support only minimum and maximum rollups.");
        var children = (task.Uid == 0 ? _document.Tasks.Where(t => t.Uid != 0) : task.Children).Where(t => t.IsNull != true && t.IsActive != false).ToArray();
        if (operation == 6) return new ProjectFormulaValue((decimal)children.Length);
        var descendants = new List<ProjectTask>(); var pending = new Stack<ProjectTask>(children.AsEnumerable().Reverse());
        while (pending.Count > 0) {
            _token.ThrowIfCancellationRequested(); var child = pending.Pop(); descendants.Add(child);
            if (descendants.Count > _maxValues) throw new InvalidDataException("Summary traversal exceeds MaxValues.");
            foreach (var nested in child.Children.Reverse()) if (nested.IsNull != true && nested.IsActive != false) pending.Push(nested);
        }
        if (operation == 2) return new ProjectFormulaValue((decimal)descendants.Count);
        var leaves = descendants.Where(t => !t.IsSummary).ToArray();
        if (operation == 7) return new ProjectFormulaValue((decimal)leaves.Length);
        var source = operation == 5 ? children : leaves;
        if (source.Length == 0) {
            if (operation == 3) return new ProjectFormulaValue(0m);
            throw new InvalidDataException("The summary has no modeled children to aggregate.");
        }
        var values = source.Select(t => Evaluate(t, identity, depth + 1)).ToArray();
        if (identity.Kind == "Flag") {
            if (operation == 0) return new ProjectFormulaValue(values.Any(v => v.Flag));
            if (operation == 1) return new ProjectFormulaValue(values.All(v => v.Flag));
            throw new NotSupportedException("Flag fields support only OR and AND rollups.");
        }
        if (operation == 0 || operation == 1) {
            var best = values[0];
            foreach (var value in values.Skip(1)) {
                int comparison = ProjectFormulaValue.Compare(value, best, _culture);
                if (operation == 0 ? comparison > 0 : comparison < 0) best = value;
            }
            return best;
        }
        decimal sum = values.Sum(v => v.NumberIn(_culture));
        return new ProjectFormulaValue(operation == 3 ? sum : sum / values.Length);
    }
}
