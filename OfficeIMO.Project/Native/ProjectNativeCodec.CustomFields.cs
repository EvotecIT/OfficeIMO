using System.Globalization;

namespace OfficeIMO.Project;

internal static partial class ProjectNativeCodec {
    // Public PjField identifiers. Only the scalar families below have a qualified native value representation.
    private static readonly uint[] TaskText = { 0x33, 0x36, 0x39, 0x3c, 0x3f, 0x42, 0x43, 0x44, 0x45, 0x46 };
    private static readonly uint[] ResourceText = { 8, 9, 0x1e, 0x1f, 0x20, 0x61, 0x62, 0x63, 0x64, 0x65 };
    private static void ReadCustomValues(ProjectDocument document, ProjectNativeRecord record, ProjectCollection<ProjectCustomFieldValue> values, bool task, CancellationToken token) {
        uint prefix = task ? 0x0b400000u : 0x0c400000u;
        void Field(uint relative, string kind, int number) {
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
            if (definition != null) definition.FieldName = kind + number.ToString(CultureInfo.InvariantCulture);
        }
        for (int number = 1; number <= 30; number++) {
            uint id = number <= 10 ? (task ? TaskText : ResourceText)[number - 1] : (task ? 0x13du : 0xe1u) + (uint)number - 11;
            Field(id, "Text", number);
        }
        for (int number = 1; number <= 20; number++) {
            Field(number <= 5 ? (task ? 0x57u : 0x70u) + (uint)number - 1 : (task ? 0x12eu : 0xcdu) + (uint)number - 6, "Number", number);
            uint flag = number <= 10 ? task ? 0x48u + (uint)number - 1 : number == 10 ? 0x7eu : 0x7fu + (uint)number - 1
                : (task ? 0x124u : 0xc3u) + (uint)number - 11;
            Field(flag, "Flag", number);
        }
        for (int number = 1; number <= 10; number++) {
            Field((task ? 0x109u : 0xadu) + (uint)number - 1, "Date", number);
            Field(number <= 3 ? (task ? 0x6au : 0x7bu) + (uint)number - 1 : (task ? 0x102u : 0xa6u) + (uint)number - 4, "Cost", number);
            Field(number <= 3 ? (task ? 0x67u : 0x75u) + (uint)number - 1 : (task ? 0x113u : 0xb7u) + (uint)number - 4, "Duration", number);
        }
    }
}
