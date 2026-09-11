using System.Globalization;
using System.Xml;

namespace OfficeIMO.Project;

internal sealed partial class ProjectNativeWriter {
    private void WriteCustomFields(IProjectNativeTableEditor editor, int uid, string parent, ProjectCollection<ProjectCustomFieldValue> values, bool task) {
        string prefix = parent + "/Custom";
        if (!ChangedTree(prefix)) return;
        Handle(prefix + "/Count");
        var catalog = task ? ProjectCustomFieldIdentity.TaskFields : ProjectCustomFieldIdentity.ResourceFields;
        var written = new HashSet<uint>();
        var retained = new HashSet<string>(values.Where(v => v.FieldId != null).Select(v => v.FieldId!));
        foreach (var old in Original(prefix).Where(p => p.Key.EndsWith("/FieldId", StringComparison.Ordinal))) {
            if (!(old.Value is string oldId) || retained.Contains(oldId) || !uint.TryParse(oldId, out uint id) || !catalog.Any(f => f.Id == id)) continue;
            editor.Set(uid, id, null);
        }
        for (int index = 0; index < values.Count; index++) {
            _token.ThrowIfCancellationRequested(); var value = values[index]; string path = prefix + "[" + index + "]";
            if (!uint.TryParse(value.FieldId, NumberStyles.None, CultureInfo.InvariantCulture, out uint id)) continue;
            var field = catalog.FirstOrDefault(f => f.Id == id); if (field == null) continue;
            if (!written.Add(id)) { AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_CUSTOM_DUPLICATE", ProjectDiagnosticSeverity.Error, "A native scalar custom field can have only one value per entity.", path)); continue; }
            // Lookup-backed fields must retain their representation until the lookup owner is implemented.
            if (value.ValueId != null || value.ValueGuid != null) continue;
            Handle(path + "/FieldId"); Handle(path + "/Value");
            if (!ChangedTree(path)) continue;
            try {
                if (value.Value == null) editor.Set(uid, id, null);
                else switch (field.Kind) {
                    case "Text": editor.Set(uid, id, Text(value.Value)); break;
                    case "Flag":
                        if (value.Value != "0" && value.Value != "1") throw new FormatException("Custom flags require 0 or 1.");
                        editor.Set(uid, id, new byte[] { value.Value == "1" ? (byte)1 : (byte)0 }); break;
                    case "Number": case "Cost": editor.Set(uid, id, Number(decimal.Parse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture))); break;
                    case "Date": editor.Set(uid, id, BitConverter.GetBytes(ProjectNativeCreation.Date(XmlConvert.ToDateTime(value.Value, XmlDateTimeSerializationMode.RoundtripKind)))); break;
                    case "Duration":
                        editor.Integer(uid, id, Exact(XmlConvert.ToTimeSpan(value.Value).Ticks / (decimal)(TimeSpan.TicksPerMinute / 10)));
                        if (field.DurationFormat != 0) {
                            int format = value.DurationFormat ?? 7;
                            int unit = format & ~32;
                            if (unit < 3 || unit > 12) throw new FormatException("Custom duration format must be a qualified working or elapsed unit.");
                            editor.Integer(uid, field.DurationFormat, format); Handle(path + "/DurationFormat");
                        }
                        break;
                }
            } catch (Exception ex) when (ex is FormatException || ex is OverflowException || ex is ArgumentException || ex is NotSupportedException) {
                AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_CUSTOM_VALUE", ProjectDiagnosticSeverity.Error, ex.Message, path));
            }
        }
        foreach (var old in Original(prefix).Where(p => !_current.ContainsKey(p.Key))) Handle(old.Key);
    }
}
