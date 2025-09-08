using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public sealed class FrontMatterBlock : IMarkdownBlock {
    private readonly List<(string Key, object? Value)> _pairs = new List<(string, object?)>();

    public static FrontMatterBlock FromObject(object data) {
        FrontMatterBlock fm = new FrontMatterBlock();
        if (data is IEnumerable<KeyValuePair<string, object?>> dict) {
            foreach (KeyValuePair<string, object?> kv in dict) fm._pairs.Add((kv.Key, kv.Value));
        } else {
            var props = data.GetType().GetProperties();
            foreach (var p in props) fm._pairs.Add((p.Name, p.GetValue(data)));
        }
        return fm;
    }

    public string Render() {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("---");
        foreach ((string k, object? v) in _pairs) {
            sb.AppendLine(k + ": " + YamlValue(v));
        }
        sb.Append("---");
        return sb.ToString();
    }

    private static string YamlValue(object? value) {
        switch (value) {
            case null:
                return string.Empty;
            case string s:
                return EscapeYamlString(s);
            case bool b:
                return b ? "true" : "false";
            case Enum e:
                return EscapeYamlString(e.ToString());
            case IEnumerable<string> ss:
                return "[" + string.Join(", ", ss.Select(EscapeYamlBareOrQuoted)) + "]";
            case System.Collections.IEnumerable ie:
                List<string> vals = new List<string>();
                foreach (object? item in ie) vals.Add(EscapeYamlBareOrQuoted(item?.ToString() ?? string.Empty));
                return "[" + string.Join(", ", vals) + "]";
            default:
                if (IsNumeric(value)) return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                return EscapeYamlString(value.ToString() ?? string.Empty);
        }
    }

    private static bool IsNumeric(object o) => o is sbyte || o is byte || o is short || o is ushort || o is int || o is uint || o is long || o is ulong || o is float || o is double || o is decimal;

    private static string EscapeYamlString(string s) {
        if (string.IsNullOrEmpty(s)) return "\"\"";
        bool needsQuotes = s.IndexOfAny(new[] { ':', '#', '\'', '"', '\n', '\r', '\t' }) >= 0 || s.Contains(' ');
        if (!needsQuotes) return s;
        // minimal escaping for quotes and backslashes
        string escaped = s.Replace("\\", "\\\\").Replace("\"", "\\\"");
        return "\"" + escaped + "\"";
    }

    private static string EscapeYamlBareOrQuoted(string s) => EscapeYamlString(s);

    public string RenderMarkdown() => Render();
    public string RenderHtml() => string.Empty;
}

