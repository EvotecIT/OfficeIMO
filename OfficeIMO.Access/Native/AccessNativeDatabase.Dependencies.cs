namespace OfficeIMO.Access;

internal sealed partial class AccessNativeDatabase {
    private void LoadDependencies() {
        var dependencies = new List<AccessDependency>();
        var objects = _document.Tables.Cast<AccessNamedObject>().Concat(_document.Queries).Concat(_document.Forms).Concat(_document.Reports).ToArray();
        void Add(Guid source, string kind, string reference, AccessNamedObject[]? targets = null) {
            var matches = (targets ?? objects).Where(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, reference)).ToArray();
            dependencies.Add(new AccessDependency(source, kind, reference, matches.Length == 1 ? matches[0].Id : (Guid?)null));
        }
        foreach (var query in _document.Queries) foreach (var reference in query.NativeRecords.Where(x => x.Attribute == 5 && x.Name1 != null))
            Add(query.Id, "query-table", reference.Name1!);
        foreach (var obj in _document.Forms.Concat(_document.Reports)) {
            if (obj.Definition == null) continue;
            string? recordSource = obj.Definition.RecordSource;
            if (recordSource != null) Add(obj.Id, "record-source", recordSource);
            var sources = _document.Tables.Where(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, recordSource)).ToArray();
            AccessNamedObject[] columns = sources.Length == 1 ? sources[0].Columns.Cast<AccessNamedObject>().ToArray() : Array.Empty<AccessNamedObject>();
            foreach (var node in Nodes(obj.Definition)) {
                if (node.ControlSource != null) Add(obj.Id, "control-source", node.ControlSource, columns);
                if (node.RowSource != null) Add(obj.Id, "row-source", node.RowSource);
            }
        }
        foreach (var relationship in _document.Relationships) foreach (var field in relationship.Fields) {
            dependencies.Add(new AccessDependency(relationship.Id, "relationship-parent", field.Parent.Name, field.Parent.Id));
            dependencies.Add(new AccessDependency(relationship.Id, "relationship-child", field.Child.Name, field.Child.Id));
        }
        _document.Dependencies = Array.AsReadOnly(dependencies.ToArray());
    }
    private static IEnumerable<AccessDesignerNode> Nodes(AccessDesignerNode root) {
        yield return root;
        foreach (var child in root.Children) foreach (var nested in Nodes(child)) yield return nested;
    }
}
