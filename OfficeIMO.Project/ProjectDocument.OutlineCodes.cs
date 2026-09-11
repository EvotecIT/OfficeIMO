using System.Globalization;

namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Resolves a task or resource outline-code selection to its masked hierarchical text; returns null when unassigned.</summary>
    public string? GetOutlineCodeText(ProjectEntity entity, string fieldId) {
        EnsureNotDisposed();
        var fields = OutlineSelections(entity, fieldId);
        var selection = fields.SingleOrDefault(v => ProjectCustomFieldIdentity.SameId(v.FieldId, fieldId));
        if (selection == null) return null;
        var index = new ProjectOutlineCodeIndex(FindOutlineTable(fieldId));
        return index.Text(index.Resolve(selection), true);
    }

    /// <summary>Assigns a hierarchical lookup value to a task or resource outline-code field, validating masks and selection restrictions before mutation.</summary>
    public void SetOutlineCodeValue(ProjectEntity entity, string fieldId, ProjectOutlineCodeLookupValue value) {
        EnsureMutable();
        if (value == null) throw new ArgumentNullException(nameof(value));
        var fields = OutlineSelections(entity, fieldId);
        var table = FindOutlineTable(fieldId);
        if (value.Document != this || !value.Attached || !table.Values.Contains(value))
            throw new ArgumentException("The selected value must belong to this field's lookup table.", nameof(value));
        var index = new ProjectOutlineCodeIndex(table); _ = index.Text(value, true);
        var existing = fields.SingleOrDefault(v => ProjectCustomFieldIdentity.SameId(v.FieldId, fieldId));
        using (BeginUpdate()) {
            var selected = existing ?? fields.Add(); selected.FieldId = fieldId;
            selected.ValueId = value.ValueId!.Value.ToString(CultureInfo.InvariantCulture); selected.ValueGuid = value.Guid;
        }
    }

    private ProjectCollection<ProjectCustomFieldValue> OutlineSelections(ProjectEntity entity, string fieldId) {
        if (entity == null) throw new ArgumentNullException(nameof(entity));
        if (fieldId == null) throw new ArgumentNullException(nameof(fieldId));
        CheckMember(entity);
        uint first = entity is ProjectTask ? 188744096u : entity is ProjectResource ? 205521174u : 0u;
        if (first == 0 || !uint.TryParse(fieldId, NumberStyles.None, CultureInfo.InvariantCulture, out uint id) || id < first || id > first + 18 || (id - first) % 2 != 0)
            throw new ArgumentException("The field must identify one of the selected entity kind's ten local outline codes.", nameof(fieldId));
        return entity is ProjectTask task ? task.OutlineCodes : ((ProjectResource)entity).OutlineCodes;
    }
    private ProjectOutlineCodeDefinition FindOutlineTable(string fieldId) {
        var field = CustomFields.SingleOrDefault(f => ProjectCustomFieldIdentity.SameId(f.FieldId, fieldId));
        var matches = OutlineCodes.Where(t => field?.LookupTableGuid != null
            ? ProjectOutlineCodeIndex.SameGuid(t.Guid, field.LookupTableGuid)
            : ProjectCustomFieldIdentity.SameId(t.FieldId, fieldId)).ToArray();
        if (matches.Length != 1) throw new InvalidDataException("The outline-code field must reference exactly one lookup table.");
        return matches[0];
    }
}
