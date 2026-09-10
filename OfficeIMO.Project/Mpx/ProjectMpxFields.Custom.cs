using System.Globalization;
using System.Xml;

namespace OfficeIMO.Project;

internal static partial class ProjectMpxFields {
    internal sealed class CustomMapping {
        internal readonly int Id;
        internal readonly string FieldId, Name, Kind;
        internal CustomMapping(int id, ProjectNativeCustomField field) { Id = id; FieldId = field.Id.ToString(CultureInfo.InvariantCulture); Name = field.Name; Kind = field.Kind; }
        internal string Parse(string text, ProjectMpxValues values) => Kind switch {
            "Text" => text, "Flag" => values.Flag(text) ? "1" : "0",
            "Number" => ProjectMpxValues.Text(values.Number(text)),
            // MSPDI and the shared model retain custom costs in hundredths of a currency unit.
            "Cost" => ProjectMpxValues.Text(checked(values.Number(text) * 100)),
            "Date" => XmlConvert.ToString(values.Date(text), XmlDateTimeSerializationMode.Unspecified),
            "Duration" => XmlConvert.ToString(TimeSpan.FromTicks(checked((long)(values.Minutes(values.Duration(text)) * TimeSpan.TicksPerMinute)))),
            _ => throw new InvalidDataException("Unknown MPX custom value kind.")
        };
        internal string Write(ProjectCustomFieldValue value, ProjectMpxValues values) {
            if (value.Value == null) return "";
            if (Kind == "Text") return value.Value;
            if (Kind == "Flag") return value.Value switch { "0" => "No", "1" => "Yes", _ => throw new InvalidDataException("MPX custom flags require 0 or 1.") };
            if (Kind == "Cost") return ProjectMpxValues.Text(decimal.Parse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture) / 100);
            if (Kind == "Date") {
                var date = XmlConvert.ToDateTime(value.Value, XmlDateTimeSerializationMode.RoundtripKind);
                if (date.Kind != DateTimeKind.Unspecified || date.Ticks % TimeSpan.TicksPerMinute != 0) throw new InvalidDataException("MPX custom dates require local wall time in whole minutes.");
                return ProjectMpxValues.Text(date);
            }
            if (Kind != "Duration") return ProjectMpxValues.Text(decimal.Parse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture));
            decimal minutes = XmlConvert.ToTimeSpan(value.Value).Ticks / (decimal)TimeSpan.TicksPerMinute;
            int format = value.DurationFormat ?? 3; int unitCode = format & ~32;
            if (unitCode < 3 || unitCode > 12) throw new InvalidDataException("Invalid custom duration format.");
            var duration = new ProjectDuration(1, (ProjectDurationUnit)((unitCode - 3) / 2), (unitCode & 1) == 0, (format & 32) != 0);
            return ProjectMpxValues.Text(new ProjectDuration(minutes / values.Minutes(duration), duration.Unit, duration.IsElapsed, duration.IsEstimated));
        }
    }
    private static readonly CustomMapping[] TaskCustom = BuildCustom(true), ResourceCustom = BuildCustom(false);
    internal static IEnumerable<CustomMapping> CustomMappings(bool task) => task ? TaskCustom : ResourceCustom;
    private static CustomMapping[] BuildCustom(bool task) {
        var result = new List<CustomMapping>();
        foreach (var field in task ? ProjectNativeCustomField.TaskFields : ProjectNativeCustomField.ResourceFields) {
            int id = field.Kind switch {
                "Text" when field.Number <= (task ? 10 : 5) => field.Number + (task ? 3 : 4),
                "Number" when task && field.Number <= 5 => field.Number + 139,
                "Flag" when task && field.Number <= 10 => field.Number + 109,
                "Cost" when task && field.Number <= 3 => field.Number + 35,
                "Duration" when task && field.Number <= 3 => field.Number + 45,
                "Date" when task && field.Name.StartsWith("Start", StringComparison.Ordinal) => field.Number <= 3 ? 58 + field.Number * 2 : 118 + field.Number * 2,
                "Date" when task && field.Name.StartsWith("Finish", StringComparison.Ordinal) => field.Number <= 3 ? 59 + field.Number * 2 : 119 + field.Number * 2,
                _ => -1
            };
            if (id >= 0) result.Add(new CustomMapping(id, field));
        }
        return result.ToArray();
    }
}
