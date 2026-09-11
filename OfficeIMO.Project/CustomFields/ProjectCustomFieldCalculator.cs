using System.Globalization;

namespace OfficeIMO.Project;

internal sealed partial class ProjectCustomFieldCalculator {
    private readonly ProjectDocument _document;
    private readonly CancellationToken _token;
    private readonly long _revision;
    private readonly int _maxOperations, _maxDepth, _maxValues;
    private readonly ProjectFormulaNesting _nesting;
    private readonly CultureInfo _culture;
    private readonly Dictionary<string, ProjectCustomFieldDefinition> _definitions = new(StringComparer.Ordinal);
    private readonly Dictionary<(ProjectEntity Entity, string Field), ProjectFormulaValue> _calculated = new();
    private readonly HashSet<(ProjectEntity Entity, string Field)> _visiting = new();
    private readonly List<ProjectCalculatedCustomField> _values = new();
    private readonly List<ProjectDiagnostic> _diagnostics = new();
    internal ProjectCustomFieldCalculator(ProjectDocument document, ProjectCustomFieldCalculationOptions options, CancellationToken token) {
        _document = document; _revision = document.Revision; _token = token;
        _maxOperations = options.MaxExpressionOperations; _maxDepth = options.MaxDepth; _maxValues = options.MaxValues;
        _nesting = new ProjectFormulaNesting(_maxDepth);
        _culture = CultureInfo.GetCultureInfo(options.CultureName ?? throw new ArgumentException("CultureName cannot be null.", nameof(options)));
    }
    internal ProjectCustomFieldCalculationResult Calculate() {
        _token.ThrowIfCancellationRequested();
        _diagnostics.AddRange(_document.Validate(_token).Diagnostics.Where(d => d.Severity == ProjectDiagnosticSeverity.Error));
        if (_diagnostics.Count > 0) return Result();
        foreach (var definition in _document.CustomFields) {
            string id = ProjectCustomFieldIdentity.NormalizeId(definition.FieldId!);
            if (_definitions.ContainsKey(id)) { Error("Custom-field definitions contain duplicate numeric identities.", "/Definition/" + id); return Result(); }
            _definitions.Add(id, definition);
        }
        foreach (var definition in _definitions.Values.Where(d => !string.IsNullOrWhiteSpace(d.Formula) || d.SummaryCalculation == 1)) {
            _token.ThrowIfCancellationRequested();
            if (definition.SummaryCalculation.HasValue && (definition.SummaryCalculation < 0 || definition.SummaryCalculation > 2)) {
                Error("The summary calculation type is outside the supported profile.", "/Definition/" + definition.FieldId); continue;
            }
            var identity = Identity(definition.FieldId!);
            if (identity == null) { Error("Only qualified scalar task and resource fields support calculation.", "/Definition/" + definition.FieldId); continue; }
            IEnumerable<ProjectEntity> entities = identity.IsTask ? _document.AllTasks.Cast<ProjectEntity>() : _document.Resources.Cast<ProjectEntity>();
            foreach (var entity in entities) {
                _token.ThrowIfCancellationRequested();
                if (entity is ProjectTask task && (task.IsNull == true || task.IsActive == false)) continue;
                if (entity is ProjectTask summary && summary.IsSummary && (definition.SummaryCalculation ?? 0) == 0) continue;
                if (entity is ProjectTask leaf && !leaf.IsSummary && string.IsNullOrWhiteSpace(definition.Formula)) continue;
                try { Evaluate(entity, identity, 0); }
                catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException || exception is FormatException ||
                    exception is OverflowException || exception is DivideByZeroException || exception is ArgumentException || exception is InvalidOperationException) {
                    Error(exception.Message, Path(entity, definition.FieldId!));
                    if (_diagnostics.Count >= 1000) return Result();
                }
            }
        }
        return Result();
    }
    private void Error(string message, string path) => _diagnostics.Add(new ProjectDiagnostic("PROJECT_CUSTOM_FIELD_CALCULATION", ProjectDiagnosticSeverity.Error, message, path));
    private ProjectCustomFieldCalculationResult Result() {
        if (_document.Revision != _revision) throw new InvalidOperationException("The document changed during custom-field calculation.");
        return new ProjectCustomFieldCalculationResult(_document, _revision, _values, _diagnostics);
    }
    private static string Path(ProjectEntity entity, string field) => (entity is ProjectTask ? "/Task" : "/Resource") + "[UID=" + entity.Uid + "]/ExtendedAttribute[FieldID=" + field + "]";
    private static ProjectCustomFieldIdentity? Identity(string id) => uint.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint numeric)
        ? ProjectCustomFieldIdentity.TaskFields.Concat(ProjectCustomFieldIdentity.ResourceFields).FirstOrDefault(f => f.Id == numeric) : null;
    private ProjectFormulaValue Evaluate(ProjectEntity entity, ProjectCustomFieldIdentity identity, int depth) {
        _token.ThrowIfCancellationRequested(); string id = identity.Id.ToString(CultureInfo.InvariantCulture); var key = (entity, id);
        if (_calculated.TryGetValue(key, out var cached)) return cached;
        if (depth >= _maxDepth) throw new InvalidDataException("Custom-field dependency depth exceeds MaxDepth.");
        if (!_visiting.Add(key)) throw new InvalidDataException("Custom-field formulas contain a dependency cycle.");
        try {
            _definitions.TryGetValue(id, out var definition);
            bool summary = entity is ProjectTask task && task.IsSummary;
            ProjectFormulaValue value;
            bool computed = false;
            if (summary && definition?.SummaryCalculation == 1) {
                value = Rollup((ProjectTask)entity, identity, definition, depth); computed = true;
            } else if (!string.IsNullOrWhiteSpace(definition?.Formula) && (!summary || definition!.SummaryCalculation == 2)) {
                value = new ProjectFormulaExpression(definition!.Formula!, name => Resolve(entity, name, depth + 1), _maxOperations, _nesting, _culture, _token).Evaluate();
                computed = true;
            } else value = ReadStored(entity, identity, definition);
            if (_calculated.Count >= _maxValues) throw new InvalidDataException("Custom-field calculation exceeds MaxValues.");
            string lexical = Format(identity, value);
            value = Parse(identity, lexical);
            _calculated.Add(key, value);
            if (computed) _values.Add(new ProjectCalculatedCustomField(identity.IsTask ? ProjectCustomFieldEntityKind.Task : ProjectCustomFieldEntityKind.Resource, entity.Uid, id, lexical));
            return value;
        } finally { _visiting.Remove(key); }
    }
    private ProjectFormulaValue Resolve(ProjectEntity entity, string name, int depth) {
        string reference = ProjectCustomFieldIdentity.NormalizeId(name.StartsWith("MSPJ", StringComparison.OrdinalIgnoreCase) ? name.Substring(4) : name);
        bool task = entity is ProjectTask;
        var catalog = task ? ProjectCustomFieldIdentity.TaskFields : ProjectCustomFieldIdentity.ResourceFields;
        var matching = catalog.Where(f => f.Id.ToString(CultureInfo.InvariantCulture) == reference || f.Name.Equals(reference, StringComparison.OrdinalIgnoreCase) ||
            _definitions.TryGetValue(f.Id.ToString(CultureInfo.InvariantCulture), out var definition) &&
            (string.Equals(definition.Alias, reference, StringComparison.OrdinalIgnoreCase) || string.Equals(definition.FieldName, reference, StringComparison.OrdinalIgnoreCase))).ToArray();
        if (matching.Length > 1) throw new InvalidDataException("The custom-field reference is ambiguous: " + name);
        if (matching.Length == 1) return Evaluate(entity, matching[0], depth);
        return Builtin(entity, reference);
    }
    private ProjectFormulaValue ReadStored(ProjectEntity entity, ProjectCustomFieldIdentity identity, ProjectCustomFieldDefinition? definition) {
        var fields = entity is ProjectTask task ? task.CustomFields : ((ProjectResource)entity).CustomFields;
        string id = identity.Id.ToString(CultureInfo.InvariantCulture);
        var stored = fields.SingleOrDefault(f => ProjectCustomFieldIdentity.SameId(f.FieldId, id));
        return Parse(identity, ProjectCustomFieldLookup.Text(_document, definition, stored));
    }
    internal static ProjectFormulaValue Parse(ProjectCustomFieldIdentity identity, string? text) => new ProjectFormulaValue(identity.Kind switch {
        "Text" => text ?? "",
        "Flag" => text == null ? false : ProjectXmlValue.ParseBool(text),
        "Date" => text == null ? throw new InvalidDataException("The referenced date field has no value.") : ProjectXmlValue.ParseDate(text),
        "Duration" => text == null ? 0m : ProjectXmlValue.ParseWork(text).Minutes,
        _ => text == null ? 0m : decimal.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture)
    });
    private string Format(ProjectCustomFieldIdentity identity, ProjectFormulaValue value) => identity.Kind switch {
        "Text" => value.TextIn(_culture),
        "Flag" => value.FlagIn(_culture) ? "1" : "0",
        "Date" => ProjectXmlValue.Date(value.Date)!,
        "Duration" => ProjectXmlValue.Work(new ProjectWork(value.NumberIn(_culture)))!,
        _ => value.NumberIn(_culture).ToString(CultureInfo.InvariantCulture)
    };
}
