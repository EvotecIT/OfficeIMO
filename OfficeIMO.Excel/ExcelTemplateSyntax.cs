using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace OfficeIMO.Excel {
    internal sealed class ExcelTemplateMarkerSyntax {
        internal ExcelTemplateMarkerSyntax(int index, int length, string text, string name, string? format) {
            Index = index;
            Length = length;
            Text = text;
            Name = name;
            Format = format;
        }

        internal int Index { get; }
        internal int Length { get; }
        internal string Text { get; }
        internal string Name { get; }
        internal string? Format { get; }
    }

    /// <summary>Lossless bounded parser for worksheet <c>{{Marker:Format}}</c> template text.</summary>
    internal sealed class ExcelTemplateSyntax {
        private ExcelTemplateSyntax(string text, IReadOnlyList<ExcelTemplateMarkerSyntax> markers) {
            Text = text;
            Markers = markers;
        }

        internal string Text { get; }
        internal IReadOnlyList<ExcelTemplateMarkerSyntax> Markers { get; }

        internal bool IsWholeMarker {
            get {
                if (Markers.Count != 1) return false;
                ExcelTemplateMarkerSyntax marker = Markers[0];
                for (int index = 0; index < marker.Index; index++) {
                    if (!char.IsWhiteSpace(Text[index])) return false;
                }
                for (int index = marker.Index + marker.Length; index < Text.Length; index++) {
                    if (!char.IsWhiteSpace(Text[index])) return false;
                }
                return true;
            }
        }

        internal static ExcelTemplateSyntax Parse(string text) {
            if (text == null) throw new ArgumentNullException(nameof(text));
            var markers = new List<ExcelTemplateMarkerSyntax>();
            int cursor = 0;
            while (cursor + 1 < text.Length) {
                int opening = text.IndexOf("{{", cursor, StringComparison.Ordinal);
                if (opening < 0) break;
                if (TryReadMarker(text, opening, out ExcelTemplateMarkerSyntax? marker)) {
                    markers.Add(marker!);
                    cursor = opening + marker!.Length;
                } else {
                    cursor = opening + 2;
                }
            }
            return new ExcelTemplateSyntax(text, new ReadOnlyCollection<ExcelTemplateMarkerSyntax>(markers));
        }

        internal string Rewrite(Func<ExcelTemplateMarkerSyntax, string> rewriter) {
            if (rewriter == null) throw new ArgumentNullException(nameof(rewriter));
            if (Markers.Count == 0) return Text;
            var output = new StringBuilder(Text.Length);
            int cursor = 0;
            foreach (ExcelTemplateMarkerSyntax marker in Markers) {
                output.Append(Text, cursor, marker.Index - cursor);
                output.Append(rewriter(marker));
                cursor = marker.Index + marker.Length;
            }
            output.Append(Text, cursor, Text.Length - cursor);
            return output.ToString();
        }

        private static bool TryReadMarker(string text, int start, out ExcelTemplateMarkerSyntax? marker) {
            marker = null;
            int cursor = start + 2;
            SkipWhitespace(text, ref cursor);
            int nameStart = cursor;
            while (cursor < text.Length && IsNameCharacter(text[cursor])) cursor++;
            if (cursor == nameStart) return false;
            string name = text.Substring(nameStart, cursor - nameStart);
            SkipWhitespace(text, ref cursor);

            string? format = null;
            if (cursor < text.Length && text[cursor] == ':') {
                cursor++;
                int formatStart = cursor;
                while (cursor + 1 < text.Length && !(text[cursor] == '}' && text[cursor + 1] == '}')) {
                    if (text[cursor] == '}') return false;
                    cursor++;
                }
                if (cursor + 1 >= text.Length) return false;
                format = text.Substring(formatStart, cursor - formatStart).Trim();
            } else {
                if (cursor + 1 >= text.Length || text[cursor] != '}' || text[cursor + 1] != '}') return false;
            }

            int end = cursor + 2;
            marker = new ExcelTemplateMarkerSyntax(start, end - start, text.Substring(start, end - start), name, format);
            return true;
        }

        private static void SkipWhitespace(string text, ref int cursor) {
            while (cursor < text.Length && char.IsWhiteSpace(text[cursor])) cursor++;
        }

        private static bool IsNameCharacter(char value) =>
            (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z')
            || (value >= '0' && value <= '9') || value == '_' || value == '.' || value == '-';
    }
}
