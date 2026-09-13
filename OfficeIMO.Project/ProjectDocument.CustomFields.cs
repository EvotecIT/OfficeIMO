namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Evaluates supported task and resource formulas against stored model values. Calculate schedules first when formula inputs depend on changed dates, work, or cost.</summary>
    public ProjectCustomFieldCalculationResult CalculateCustomFields(ProjectCustomFieldCalculationOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (_batchDepth != 0) throw new InvalidOperationException("Finish the update scope before calculating custom fields.");
        options ??= new ProjectCustomFieldCalculationOptions();
        if (options.MaxDepth < 1 || options.MaxDepth > 128 || options.MaxExpressionOperations < 1 || options.MaxExpressionOperations > 1000000 || options.MaxValues < 1)
            throw new ArgumentOutOfRangeException(nameof(options));
        return new ProjectCustomFieldCalculator(this, options, cancellationToken).Calculate();
    }
    /// <summary>Atomically applies a complete custom-field proposal to its unchanged originating document.</summary>
    public void ApplyCustomFields(ProjectCustomFieldCalculationResult result, CancellationToken cancellationToken = default) {
        EnsureMutable();
        if (result == null) throw new ArgumentNullException(nameof(result));
        result.Report.ThrowIfErrors();
        if (result.Document != this || result.ModelRevision != Revision || _batchDepth != 0)
            throw new InvalidOperationException("The custom-field proposal belongs to another document or a stale/in-progress revision.");
        var updates = new List<(ProjectCollection<ProjectCustomFieldValue> Fields, ProjectCustomFieldValue? Existing, ProjectCalculatedCustomField Value)>();
        foreach (var value in result.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            var fields = value.EntityKind == ProjectCustomFieldEntityKind.Task ? Tasks.GetByUid(value.EntityUid).CustomFields : Resources.GetByUid(value.EntityUid).CustomFields;
            updates.Add((fields, fields.SingleOrDefault(f => ProjectCustomFieldIdentity.SameId(f.FieldId, value.FieldId)), value));
        }
        cancellationToken.ThrowIfCancellationRequested();
        using (BeginUpdate()) foreach (var update in updates) {
            var field = update.Existing ?? update.Fields.Add();
            field.FieldId = update.Value.FieldId; field.Value = update.Value.Value; field.ValueId = null; field.ValueGuid = null;
        }
    }
}
