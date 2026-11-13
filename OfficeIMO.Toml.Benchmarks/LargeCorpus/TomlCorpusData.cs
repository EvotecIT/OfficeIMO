using System.Globalization;
using System.Text;

namespace OfficeIMO.Toml.Benchmarks.LargeCorpus;

/// <summary>
/// Canonical text of a loaded document's data, for comparing OfficeIMO with Tomlyn and a document with its edited and
/// reloaded copy. Keys are sorted; values carry their type, and date-times their kind and exact fields.
/// </summary>
internal static class TomlCorpusData {
    public static string Of(TomlNode node) {
        var text = new StringBuilder();
        Write(text, node);
        return text.ToString();
    }

    public static string OfTomlyn(object value) {
        var text = new StringBuilder();
        WriteTomlyn(text, value);
        return text.ToString();
    }

    private static void Write(StringBuilder text, TomlNode node) {
        switch (node) {
            case TomlTable table:
                text.Append('{');
                foreach (KeyValuePair<string, TomlNode> entry in table.OrderBy(e => e.Key, StringComparer.Ordinal)) {
                    Key(text, entry.Key);
                    Write(text, entry.Value);
                    text.Append(',');
                }
                text.Append('}');
                break;
            case TomlTableArray tables:
                text.Append('[');
                foreach (TomlTable table in tables) {
                    Write(text, table);
                    text.Append(',');
                }
                text.Append(']');
                break;
            case TomlArray array:
                text.Append('[');
                foreach (TomlNode item in array) {
                    Write(text, item);
                    text.Append(',');
                }
                text.Append(']');
                break;
            case TomlValue value:
                Scalar(text, value.Value);
                break;
        }
    }

    private static void WriteTomlyn(StringBuilder text, object value) {
        switch (value) {
            case Tomlyn.Model.TomlTable table:
                text.Append('{');
                foreach (KeyValuePair<string, object> entry in table.OrderBy(e => e.Key, StringComparer.Ordinal)) {
                    Key(text, entry.Key);
                    WriteTomlyn(text, entry.Value);
                    text.Append(',');
                }
                text.Append('}');
                break;
            case Tomlyn.Model.TomlTableArray tables:
                text.Append('[');
                foreach (Tomlyn.Model.TomlTable table in tables) {
                    WriteTomlyn(text, table);
                    text.Append(',');
                }
                text.Append(']');
                break;
            case Tomlyn.Model.TomlArray array:
                text.Append('[');
                foreach (object? item in array) {
                    WriteTomlyn(text, item!);
                    text.Append(',');
                }
                text.Append(']');
                break;
            case Tomlyn.TomlDateTime dateTime:
                DateTimeOffset moment = dateTime.DateTime;
                switch (dateTime.Kind) {
                    case Tomlyn.TomlDateTimeKind.LocalDate:
                        text.Append("date:").Append(moment.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                        break;
                    case Tomlyn.TomlDateTimeKind.LocalTime:
                        text.Append("time:").Append(moment.TimeOfDay.Ticks.ToString(CultureInfo.InvariantCulture));
                        break;
                    case Tomlyn.TomlDateTimeKind.LocalDateTime:
                        text.Append("local:").Append(moment.DateTime.Ticks.ToString(CultureInfo.InvariantCulture));
                        break;
                    default:
                        text.Append("offset:").Append(moment.UtcTicks.ToString(CultureInfo.InvariantCulture));
                        break;
                }
                break;
            default:
                Scalar(text, value);
                break;
        }
    }

    private static void Key(StringBuilder text, string key) => text.Append('"').Append(key.Replace("\\", "\\\\").Replace("\"", "\\\"")).Append("\":");

    private static void Scalar(StringBuilder text, object value) {
        switch (value) {
            case string s:
                text.Append("s\"").Append(s.Replace("\\", "\\\\").Replace("\"", "\\\"")).Append('"');
                break;
            case long l:
                text.Append("i:").Append(l.ToString(CultureInfo.InvariantCulture));
                break;
            case double d:
                text.Append("f:").Append(double.IsNaN(d) ? "nan" : d.ToString("R", CultureInfo.InvariantCulture));
                break;
            case bool b:
                text.Append(b ? "true" : "false");
                break;
            case TomlDateTime dateTime:
                switch (dateTime.Kind) {
                    case TomlDateTimeKind.LocalDate:
                        text.Append("date:").Append(dateTime.ToDateTime().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                        break;
                    case TomlDateTimeKind.LocalTime:
                        text.Append("time:").Append(Ticks(dateTime.ToTimeOfDay().Ticks, dateTime));
                        break;
                    case TomlDateTimeKind.LocalDateTime:
                        text.Append("local:").Append(Ticks(dateTime.ToDateTime().Ticks, dateTime));
                        break;
                    default:
                        text.Append("offset:").Append(Ticks(dateTime.ToDateTimeOffset().UtcTicks, dateTime));
                        break;
                }
                break;
            default:
                text.Append("?").Append(value);
                break;
        }
    }

    /// <summary>Ticks truncated to 100 ns, as Tomlyn keeps them.</summary>
    private static string Ticks(long ticks, TomlDateTime dateTime) =>
        (ticks - ticks % TimeSpan.TicksPerSecond + dateTime.Nanosecond / 100).ToString(CultureInfo.InvariantCulture);
}
