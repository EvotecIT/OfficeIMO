using System.Text;

namespace OfficeIMO.PowerPoint.LegacyPpt.Capabilities {
    public static partial class LegacyPptCapabilityCatalog {
        /// <summary>Serializes the complete capability contract as deterministic JSON.</summary>
        public static string ToJson(bool indented = true) {
            string newline = indented ? Environment.NewLine : string.Empty;
            string i1 = indented ? "  " : string.Empty;
            string i2 = indented ? "    " : string.Empty;
            string i3 = indented ? "      " : string.Empty;
            var json = new StringBuilder();
            json.Append('{').Append(newline)
                .Append(i1).Append("\"schemaVersion\":").Append(SchemaVersion).Append(',').Append(newline)
                .Append(i1).Append("\"hasRemainingParityWork\":")
                .Append(HasRemainingParityWork ? "true" : "false").Append(',').Append(newline)
                .Append(i1).Append("\"capabilities\": [").Append(newline);

            for (int index = 0; index < CapabilityRows.Count; index++) {
                LegacyPptCapability row = CapabilityRows[index];
                json.Append(i2).Append('{').Append(newline)
                    .Append(i3).Append("\"feature\":\"").Append(EscapeJson(row.Feature.ToString())).Append("\",").Append(newline)
                    .Append(i3).Append("\"category\":\"").Append(EscapeJson(row.Category)).Append("\",").Append(newline)
                    .Append(i3).Append("\"description\":\"").Append(EscapeJson(row.Description)).Append("\",").Append(newline)
                    .Append(i3).Append("\"representability\":\"").Append(row.Representability).Append("\",").Append(newline)
                    .Append(i3).Append("\"importToEditableModel\":\"").Append(row.ImportToEditableModel).Append("\",").Append(newline)
                    .Append(i3).Append("\"newBinaryWrite\":\"").Append(row.NewBinaryWrite).Append("\",").Append(newline)
                    .Append(i3).Append("\"binaryRoundTrip\":\"").Append(row.BinaryRoundTrip).Append("\",").Append(newline)
                    .Append(i3).Append("\"pptxToBinary\":\"").Append(row.PptxToBinary).Append("\",").Append(newline)
                    .Append(i3).Append("\"note\":\"").Append(EscapeJson(row.Note)).Append("\"").Append(newline)
                    .Append(i2).Append('}');
                if (index + 1 < CapabilityRows.Count) json.Append(',');
                json.Append(newline);
            }

            json.Append(i1).Append(']').Append(newline).Append('}');
            return json.ToString();
        }

        /// <summary>Formats the capability contract as a human-readable Markdown table.</summary>
        public static string ToMarkdown() {
            var markdown = new StringBuilder();
            markdown.AppendLine("# Binary PowerPoint capability contract");
            markdown.AppendLine();
            markdown.AppendLine($"Schema version: {SchemaVersion}");
            markdown.AppendLine();
            markdown.AppendLine("| Category | Feature | Representation | Import | New binary | Binary round-trip | PPTX to binary | Note |");
            markdown.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- |");
            foreach (LegacyPptCapability row in CapabilityRows) {
                markdown.Append("| ").Append(EscapeMarkdown(row.Category))
                    .Append(" | ").Append(row.Feature)
                    .Append(" | ").Append(row.Representability)
                    .Append(" | ").Append(row.ImportToEditableModel)
                    .Append(" | ").Append(row.NewBinaryWrite)
                    .Append(" | ").Append(row.BinaryRoundTrip)
                    .Append(" | ").Append(row.PptxToBinary)
                    .Append(" | ").Append(EscapeMarkdown(row.Note)).AppendLine(" |");
            }
            return markdown.ToString();
        }

        private static string EscapeJson(string value) => (value ?? string.Empty)
            .Replace("\\", "\\\\").Replace("\"", "\\\"")
            .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");

        private static string EscapeMarkdown(string value) => (value ?? string.Empty)
            .Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
    }
}
