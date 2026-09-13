using System.Globalization;

namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Returns a field's shared lookup table selected by Ltuid; null means the field uses no shared table.</summary>
    public ProjectOutlineCodeDefinition? GetCustomFieldLookupTable(ProjectCustomFieldDefinition definition) {
        EnsureNotDisposed();
        if (definition == null) throw new ArgumentNullException(nameof(definition));
        CheckMember(definition);
        if (!CustomFields.Contains(definition)) throw new ArgumentException("The definition must belong to this document.", nameof(definition));
        if (definition.LookupTableGuid == null) return null;
        var matches = OutlineCodes.Where(t => ProjectOutlineCodeIndex.SameGuid(t.Guid, definition.LookupTableGuid)).ToArray();
        if (matches.Length != 1) throw new InvalidDataException("The field must reference exactly one shared lookup table.");
        return matches[0];
    }

    /// <summary>Assigns a flat shared lookup entry to a scalar task or resource field, retaining its value and identity.</summary>
    public void SetCustomFieldLookupValue(ProjectEntity entity, ProjectCustomFieldDefinition definition, ProjectOutlineCodeLookupValue lookup) {
        EnsureMutable();
        if (entity == null) throw new ArgumentNullException(nameof(entity));
        if (lookup == null) throw new ArgumentNullException(nameof(lookup));
        CheckMember(entity); CheckMember(lookup);
        var table = GetCustomFieldLookupTable(definition);
        if (table == null || !table.Values.Contains(lookup)) throw new ArgumentException("The entry must belong to the field's shared lookup table.", nameof(lookup));
        if (lookup.ParentValueId > 0) throw new NotSupportedException("Scalar custom fields support flat shared lookup values.");
        if (!string.IsNullOrWhiteSpace(definition.Formula)) throw new InvalidOperationException("A formula-owned field cannot be assigned a lookup value.");
        var catalog = entity is ProjectTask ? ProjectCustomFieldIdentity.TaskFields : entity is ProjectResource ? ProjectCustomFieldIdentity.ResourceFields : null;
        var identity = catalog?.SingleOrDefault(f => f.Id.ToString(CultureInfo.InvariantCulture) == ProjectCustomFieldIdentity.NormalizeId(definition.FieldId ?? ""));
        if (identity == null) throw new ArgumentException("The field identity does not belong to the selected entity kind.", nameof(definition));
        _ = new ProjectOutlineCodeIndex(table);
        if (lookup.Value == null) throw new InvalidDataException("A scalar lookup entry requires a value.");
        _ = ProjectCustomFieldCalculator.Parse(identity, lookup.Value);
        var fields = entity is ProjectTask task ? task.CustomFields : ((ProjectResource)entity).CustomFields;
        var existing = fields.SingleOrDefault(f => ProjectCustomFieldIdentity.SameId(f.FieldId, definition.FieldId));
        using (BeginUpdate()) {
            var value = existing ?? fields.Add(); value.FieldId = definition.FieldId; value.Value = lookup.Value;
            value.ValueId = lookup.ValueId!.Value.ToString(CultureInfo.InvariantCulture); value.ValueGuid = lookup.Guid;
        }
    }
    /// <summary>Assigns a modeled lookup entry to a qualified scalar task or resource field, retaining its value and lookup identity.</summary>
    public void SetCustomFieldLookupValue(ProjectEntity entity, ProjectCustomFieldDefinition definition, ProjectLookupValue lookup) {
        EnsureMutable();
        if (entity == null) throw new ArgumentNullException(nameof(entity));
        if (definition == null) throw new ArgumentNullException(nameof(definition));
        if (lookup == null) throw new ArgumentNullException(nameof(lookup));
        entity.EnsureAttached(); definition.EnsureAttached(); lookup.EnsureAttached();
        if (entity.Document != this || definition.Document != this || lookup.Document != this || !CustomFields.Contains(definition) || !definition.LookupValues.Contains(lookup))
            throw new ArgumentException("Entity, definition, and lookup entry must be attached to this document and the selected lookup list.");
        if (definition.LookupTableGuid != null)
            throw new InvalidOperationException("This field uses a shared lookup table. Select an entry from GetCustomFieldLookupTable instead of its legacy inline list.");
        bool task = entity is ProjectTask;
        if (!task && !(entity is ProjectResource)) throw new ArgumentException("Only task and resource custom fields are supported.", nameof(entity));
        var catalog = task ? ProjectCustomFieldIdentity.TaskFields : ProjectCustomFieldIdentity.ResourceFields;
        if (!uint.TryParse(definition.FieldId, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint id) || !catalog.Any(f => f.Id == id))
            throw new ArgumentException("The field identity does not belong to the selected entity kind.", nameof(definition));
        if (!string.IsNullOrWhiteSpace(definition.Formula)) throw new InvalidOperationException("A formula-owned field cannot be assigned a lookup value.");
        if (lookup.Value == null || !lookup.Id.HasValue || definition.LookupValues.Count(v => v.Id == lookup.Id) != 1 ||
            lookup.Guid != null && definition.LookupValues.Count(v => string.Equals(v.Guid, lookup.Guid, StringComparison.OrdinalIgnoreCase)) != 1)
            throw new InvalidOperationException("A lookup assignment requires a value and a unique entry identity.");
        _ = ProjectCustomFieldCalculator.Parse(catalog.Single(f => f.Id == id), lookup.Value);
        var fields = task ? ((ProjectTask)entity).CustomFields : ((ProjectResource)entity).CustomFields;
        var existing = fields.SingleOrDefault(f => ProjectCustomFieldIdentity.SameId(f.FieldId, definition.FieldId));
        using (BeginUpdate()) {
            var value = existing ?? fields.Add(); value.FieldId = definition.FieldId; value.Value = lookup.Value;
            value.ValueId = lookup.Id.Value.ToString(CultureInfo.InvariantCulture); value.ValueGuid = lookup.Guid;
        }
    }
}
