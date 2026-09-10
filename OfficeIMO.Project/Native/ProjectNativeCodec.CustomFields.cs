using System.Globalization;

namespace OfficeIMO.Project;

internal static partial class ProjectNativeCodec {
    private static void ReadCustomValues(ProjectDocument document, ProjectNativeRecord record, ProjectCollection<ProjectCustomFieldValue> values, bool task, CancellationToken token) {
        uint prefix = task ? 0x0b400000u : 0x0c400000u;
        void Field(uint relative, string kind, int number, string name) {
            token.ThrowIfCancellationRequested();
            uint id = prefix | relative;
            string fieldId = id.ToString(CultureInfo.InvariantCulture);
            string? text = null; int? durationFormat = null;
            if (kind == "Flag") {
                bool? flag = record.Boolean(id);
                if (flag == true || (flag.HasValue && document.CustomFields.Any(f => f.FieldId == fieldId))) text = flag == true ? "1" : "0";
            } else if (record.Value(id).HasValue) {
                switch (kind) {
                    case "Text": text = record.Text(id); break;
                    case "Number": case "Cost": text = ProjectXmlValue.Number(record.Number(id)); break;
                    case "Date": text = ProjectXmlValue.Date(record.Date(id)); break;
                    case "Duration":
                        text = ProjectXmlValue.Span(ProjectXmlValue.MinutesToSpan((record.Integer(id) ?? 0) / 10m));
                        if (task) durationFormat = record.Integer(prefix | (number <= 3 ? 0xb7u + (uint)number - 1 : 0x151u + (uint)number - 4));
                        break;
                }
            }
            if (text == null) return;
            var value = values.Add(); value.FieldId = fieldId; value.Value = text; value.DurationFormat = durationFormat;
            var definition = document.CustomFields.FirstOrDefault(f => f.FieldId == fieldId);
            if (definition != null) definition.FieldName = name;
        }
        foreach (var field in task ? ProjectNativeCustomField.TaskFields : ProjectNativeCustomField.ResourceFields) Field(field.Relative, field.Kind, field.Number, field.Name);
    }
}
