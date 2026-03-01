using System.Globalization;

namespace OfficeIMO.Reader;

internal static class ReaderCapabilityManifestJson {
    public static string Serialize(ReaderCapabilityManifest manifest, bool indented) {
        if (manifest == null) throw new ArgumentNullException(nameof(manifest));

        var handlers = manifest.Handlers ?? Array.Empty<ReaderHandlerCapability>();
        var sb = new StringBuilder(8_192);
        int depth = 0;

        void AppendNewLine() {
            if (indented) {
                sb.AppendLine();
            }
        }

        void AppendIndent() {
            if (!indented) return;
            for (int i = 0; i < depth; i++) {
                sb.Append("  ");
            }
        }

        sb.Append('{');
        AppendNewLine();
        depth++;
        AppendIndent();
        sb.Append("\"schemaId\":\"");
        sb.Append(Escape(manifest.SchemaId ?? ReaderCapabilitySchema.Id));
        sb.Append("\",");
        AppendNewLine();

        AppendIndent();
        sb.Append("\"schemaVersion\":");
        sb.Append(manifest.SchemaVersion.ToString(CultureInfo.InvariantCulture));
        sb.Append(',');
        AppendNewLine();

        AppendIndent();
        sb.Append("\"handlers\":[");
        AppendNewLine();

        depth++;
        for (int i = 0; i < handlers.Count; i++) {
            var handler = handlers[i];
            AppendIndent();
            sb.Append('{');
            AppendNewLine();

            depth++;
            WriteString("id", handler.Id, trailingComma: true);
            WriteString("displayName", handler.DisplayName, trailingComma: true);
            WriteNullableString("description", handler.Description, trailingComma: true);
            WriteString("kind", handler.Kind.ToString(), trailingComma: true);
            WriteStringArray("extensions", handler.Extensions ?? Array.Empty<string>(), trailingComma: true);
            WriteBoolean("isBuiltIn", handler.IsBuiltIn, trailingComma: true);
            WriteBoolean("supportsPath", handler.SupportsPath, trailingComma: true);
            WriteBoolean("supportsStream", handler.SupportsStream, trailingComma: true);
            WriteString("schemaId", handler.SchemaId ?? ReaderCapabilitySchema.Id, trailingComma: true);
            WriteNumber("schemaVersion", handler.SchemaVersion, trailingComma: true);
            WriteNullableNumber("defaultMaxInputBytes", handler.DefaultMaxInputBytes, trailingComma: true);
            WriteString("warningBehavior", handler.WarningBehavior.ToString(), trailingComma: true);
            WriteBoolean("deterministicOutput", handler.DeterministicOutput, trailingComma: false);
            depth--;

            AppendNewLine();
            AppendIndent();
            sb.Append('}');
            if (i < handlers.Count - 1) {
                sb.Append(',');
            }
            AppendNewLine();
        }

        depth--;
        AppendIndent();
        sb.Append(']');
        AppendNewLine();
        depth--;
        AppendIndent();
        sb.Append('}');

        return sb.ToString();

        void WritePropertyName(string name) {
            AppendIndent();
            sb.Append('"');
            sb.Append(name);
            sb.Append("\":");
        }

        void WriteString(string name, string? value, bool trailingComma) {
            WritePropertyName(name);
            sb.Append('"');
            sb.Append(Escape(value ?? string.Empty));
            sb.Append('"');
            if (trailingComma) sb.Append(',');
            AppendNewLine();
        }

        void WriteNullableString(string name, string? value, bool trailingComma) {
            WritePropertyName(name);
            if (value == null) {
                sb.Append("null");
            } else {
                sb.Append('"');
                sb.Append(Escape(value));
                sb.Append('"');
            }
            if (trailingComma) sb.Append(',');
            AppendNewLine();
        }

        void WriteStringArray(string name, IReadOnlyList<string> values, bool trailingComma) {
            WritePropertyName(name);
            sb.Append('[');
            for (int i = 0; i < values.Count; i++) {
                if (i > 0) sb.Append(',');
                sb.Append('"');
                sb.Append(Escape(values[i] ?? string.Empty));
                sb.Append('"');
            }
            sb.Append(']');
            if (trailingComma) sb.Append(',');
            AppendNewLine();
        }

        void WriteBoolean(string name, bool value, bool trailingComma) {
            WritePropertyName(name);
            sb.Append(value ? "true" : "false");
            if (trailingComma) sb.Append(',');
            AppendNewLine();
        }

        void WriteNumber(string name, int value, bool trailingComma) {
            WritePropertyName(name);
            sb.Append(value.ToString(CultureInfo.InvariantCulture));
            if (trailingComma) sb.Append(',');
            AppendNewLine();
        }

        void WriteNullableNumber(string name, long? value, bool trailingComma) {
            WritePropertyName(name);
            if (value.HasValue) {
                sb.Append(value.Value.ToString(CultureInfo.InvariantCulture));
            } else {
                sb.Append("null");
            }
            if (trailingComma) sb.Append(',');
            AppendNewLine();
        }
    }

    private static string Escape(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;

        var sb = new StringBuilder(value.Length + 8);
        foreach (var ch in value) {
            switch (ch) {
                case '"':
                    sb.Append("\\\"");
                    break;
                case '\\':
                    sb.Append("\\\\");
                    break;
                case '\b':
                    sb.Append("\\b");
                    break;
                case '\f':
                    sb.Append("\\f");
                    break;
                case '\n':
                    sb.Append("\\n");
                    break;
                case '\r':
                    sb.Append("\\r");
                    break;
                case '\t':
                    sb.Append("\\t");
                    break;
                default:
                    if (char.IsControl(ch)) {
                        sb.Append("\\u");
                        sb.Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                    } else {
                        sb.Append(ch);
                    }
                    break;
            }
        }

        return sb.ToString();
    }
}
