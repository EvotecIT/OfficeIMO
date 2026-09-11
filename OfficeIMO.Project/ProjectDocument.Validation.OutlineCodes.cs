namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    private void CheckOutlineCodes(Finding add, CancellationToken token) {
        var guids = new HashSet<Guid>();
        var indexes = new Dictionary<ProjectOutlineCodeDefinition, ProjectOutlineCodeIndex>();
        var sharedFields = new Dictionary<string, ProjectCustomFieldDefinition>(StringComparer.Ordinal);
        foreach (var table in OutlineCodes) {
            token.ThrowIfCancellationRequested();
            if (table.Guid != null && (!System.Guid.TryParse(table.Guid, out var guid) || !guids.Add(guid)))
                add("PROJECT_OUTLINE_IDENTITY", "Outline-code table GUIDs must be valid and unique.", "/OutlineCodes");
            try { indexes.Add(table, new ProjectOutlineCodeIndex(table, token)); }
            catch (InvalidDataException exception) { add("PROJECT_OUTLINE_HIERARCHY", exception.Message, "/OutlineCodes"); }
        }
        foreach (var definition in CustomFields) {
            token.ThrowIfCancellationRequested();
            if (definition.LookupTableGuid == null) continue;
            string id = ProjectCustomFieldIdentity.NormalizeId(definition.FieldId ?? "");
            try { _ = GetCustomFieldLookupTable(definition); }
            catch (InvalidDataException exception) { add("PROJECT_LOOKUP_TABLE_REFERENCE", exception.Message, "/Definition/" + id); }
            if (sharedFields.ContainsKey(id))
                add("PROJECT_LOOKUP_TABLE_REFERENCE", "Shared lookup fields require unique numeric field identities.", "/Definition/" + id);
            else sharedFields.Add(id, definition);
        }
        void CheckScalar(ProjectEntity entity, ProjectCollection<ProjectCustomFieldValue> fields) {
            foreach (var selection in fields) {
                token.ThrowIfCancellationRequested();
                if (selection.FieldId == null || !sharedFields.TryGetValue(ProjectCustomFieldIdentity.NormalizeId(selection.FieldId), out var definition)) continue;
                try { _ = ProjectCustomFieldLookup.Text(this, definition, selection); }
                catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException) {
                    string path = "/" + (entity is ProjectTask ? "Task" : "Resource") + "[UID=" + entity.Uid + "]/ExtendedAttribute";
                    add("PROJECT_LOOKUP_VALUE_REFERENCE", exception.Message, path);
                }
            }
        }
        void Check(ProjectEntity entity, ProjectCollection<ProjectCustomFieldValue> fields) {
            var ids = new HashSet<string>();
            foreach (var selection in fields) {
                token.ThrowIfCancellationRequested();
                string path = "/" + (entity is ProjectTask ? "Task" : "Resource") + "[UID=" + entity.Uid + "]/OutlineCode";
                if (selection.FieldId == null || !ids.Add(ProjectCustomFieldIdentity.NormalizeId(selection.FieldId))) {
                    add("PROJECT_OUTLINE_SELECTION", "Outline-code selections require unique field identities.", path); continue;
                }
                if (selection.Value != null || selection.DurationFormat != null) {
                    add("PROJECT_OUTLINE_SELECTION", "Outline-code selections store value identities, not scalar values or durations.", path); continue;
                }
                try {
                    _ = OutlineSelections(entity, selection.FieldId);
                    var table = FindOutlineTable(selection.FieldId);
                    if (indexes.TryGetValue(table, out var index)) _ = index.Text(index.Resolve(selection), true);
                }
                catch (Exception exception) when (exception is InvalidDataException || exception is ArgumentException || exception is NotSupportedException || exception is InvalidOperationException) {
                    add("PROJECT_OUTLINE_SELECTION", exception.Message, path);
                }
            }
        }
        foreach (var task in AllTasks) { Check(task, task.OutlineCodes); CheckScalar(task, task.CustomFields); }
        foreach (var resource in Resources) { Check(resource, resource.OutlineCodes); CheckScalar(resource, resource.CustomFields); }
    }
}
