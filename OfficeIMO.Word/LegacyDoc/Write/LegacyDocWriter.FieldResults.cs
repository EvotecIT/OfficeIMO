using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        /// <summary>Writes a cached field without evaluating it or collapsing its result runs.</summary>
        private static void AppendSupportedField(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            string instruction,
            LegacyDocWritableFieldResult result,
            string emptyResultText,
            int characterOffset) {
            AppendFormattedNoteText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter, characterOffset);
            AppendFormattedNoteText(text, runs, instruction, LegacyDocWritableFormatting.Plain, characterOffset);
            AppendFormattedNoteText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter, characterOffset);
            int resultStartCharacter = characterOffset + text.Length;
            string resultText = result.Text.ToString();
            if (resultText.Length == 0) {
                AppendFormattedNoteText(text, runs, emptyResultText, result.EmptyFormatting, characterOffset);
            } else {
                int position = 0;
                foreach (LegacyDocWritableRun run in result.Runs) {
                    if (run.StartCharacter > position) {
                        AppendFormattedNoteText(text, runs, resultText.Substring(position, run.StartCharacter - position), LegacyDocWritableFormatting.Plain, characterOffset);
                    }
                    AppendFormattedNoteText(text, runs, resultText.Substring(run.StartCharacter, run.Length), run.Formatting, characterOffset);
                    position = run.EndCharacter;
                }
                if (position < resultText.Length) {
                    AppendFormattedNoteText(text, runs, resultText.Substring(position), LegacyDocWritableFormatting.Plain, characterOffset);
                }
            }

            foreach (LegacyDocFieldBookmarkMarker marker in result.BookmarkMarkers) {
                int characterPosition = resultStartCharacter + marker.ResultOffset;
                if (marker.BookmarkStart.Exists) {
                    bookmarks.AddStart(marker.BookmarkStart, characterPosition);
                } else if (marker.BookmarkEnd.Exists) {
                    bookmarks.AddEnd(marker.BookmarkEnd, characterPosition);
                }
            }
            int endCharacter = characterOffset + text.Length;
            text.Append(LegacyDocField.End);
            runs.Add(new LegacyDocWritableRun(endCharacter, 1, LegacyDocWritableFormatting.SpecialCharacter, fieldLocked: result.FieldLocked));
        }

        private static LegacyDocWritableFieldResult ReadSimpleFieldResult(
            LegacyDocElement field,
            bool allowHyperlinkRunStyle = false,
            LegacyDocWritableFormatting inheritedFormatting = default) {
            var result = new LegacyDocWritableFieldResult(inheritedFormatting, allowHyperlinkRunStyle) { FieldLocked = field.Attr("fldLock") is { } fieldLock && LegacyDocValues.ParseOnOff(fieldLock) };
            foreach (LegacyDocElement child in field.Elements()) {
                result.AppendChild(child);
            }
            return result;
        }

        private static void AppendSupportedFieldResultElementText(StringBuilder result, LegacyDocElement element) {
            switch (element.WName) {
                case "t":
                    result.Append(element.Text);
                    break;
                case "delText":
                    result.Append(element.Text);
                    break;
                case "tab":
                    result.Append('\t');
                    break;
                case "cr":
                    result.Append(LegacyDocSpecialCharacters.TextWrappingBreak);
                    break;
                case "noBreakHyphen":
                    result.Append(LegacyDocSpecialCharacters.NoBreakHyphen);
                    break;
                case "softHyphen":
                    result.Append(LegacyDocSpecialCharacters.SoftHyphen);
                    break;
                case "br":
                    result.Append(GetSupportedFieldResultBreakCharacter(element));
                    break;
                default:
                    throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} field result runs only when they contain text, tabs, carriage returns, soft/no-break hyphens, and supported breaks. Unsupported field result element: {element.LocalName}.");
            }
        }

        private static char GetSupportedFieldResultBreakCharacter(LegacyDocElement breakNode) {
            string? breakType = breakNode.Attr("type");
            if (breakType == null || breakType == "textWrapping") {
                return LegacyDocSpecialCharacters.TextWrappingBreak;
            }
            if (breakType == "page") {
                return LegacyDocSpecialCharacters.PageBreak;
            }
            if (breakType == "column") {
                return LegacyDocSpecialCharacters.ColumnBreak;
            }
            throw new NotSupportedException($"Native DOC saving currently supports text-wrapping, page, and column breaks in {SupportedFieldNames} field result runs only. Unsupported break type: {breakType}.");
        }

        /// <summary>Result-relative text, native run encodings, and zero-width bookmark markers shared by all stories.</summary>
        private sealed class LegacyDocWritableFieldResult {
            private readonly LegacyDocWritableFormatting _inheritedFormatting;
            private readonly bool _allowHyperlinkRunStyle;
            private LegacyDocWritableFormatting? _emptyFormatting;

            internal LegacyDocWritableFieldResult(LegacyDocWritableFormatting inheritedFormatting = default, bool allowHyperlinkRunStyle = false) {
                _inheritedFormatting = inheritedFormatting;
                _allowHyperlinkRunStyle = allowHyperlinkRunStyle;
            }

            internal bool FieldLocked { get; set; }
            internal StringBuilder Text { get; } = new();
            internal List<LegacyDocWritableRun> Runs { get; } = new();
            internal List<LegacyDocFieldBookmarkMarker> BookmarkMarkers { get; } = new();
            internal LegacyDocWritableFormatting EmptyFormatting => _emptyFormatting ?? _inheritedFormatting;

            internal void AppendChild(LegacyDocElement child, LegacyDocRevision revision = default) {
                switch (child.WName) {
                    case "r":
                        LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(child.Child("rPr"), _allowHyperlinkRunStyle)
                            .WithInheritedFormatting(_inheritedFormatting);
                        if (revision.HasValue) formatting = formatting.WithRevision(revision);
                        _emptyFormatting ??= formatting;
                        var runText = new StringBuilder();
                        foreach (LegacyDocElement element in child.Elements()) {
                            if (element.Is("rPr") || element.Is("lastRenderedPageBreak")) continue;
                            AppendSupportedFieldResultElementText(runText, element);
                        }
                        AppendText(runText.ToString(), formatting);
                        break;
                    case "bookmarkStart":
                        BookmarkMarkers.Add(new LegacyDocFieldBookmarkMarker(child, LegacyDocElement.None, Text.Length));
                        break;
                    case "bookmarkEnd":
                        BookmarkMarkers.Add(new LegacyDocFieldBookmarkMarker(LegacyDocElement.None, child, Text.Length));
                        break;
                    case "ins" when !revision.HasValue && !_inheritedFormatting.Revision.HasValue:
                        AppendRevision(child, LegacyDocRevisionKind.Inserted);
                        break;
                    case "del" when !revision.HasValue && !_inheritedFormatting.Revision.HasValue:
                        AppendRevision(child, LegacyDocRevisionKind.Deleted);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) return;
                        throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} field results only with text runs, bookmarks, and basic tracked insertions or deletions. Nested fields and nested revisions are not supported. Unsupported field result element: {child.LocalName}.");
                }
            }

            internal void AppendElement(LegacyDocElement element, LegacyDocWritableFormatting formatting) {
                var value = new StringBuilder();
                AppendSupportedFieldResultElementText(value, element);
                AppendText(value.ToString(), formatting);
            }

            internal void AppendText(string value, LegacyDocWritableFormatting formatting) {
                _emptyFormatting ??= formatting;
                AppendFormattedText(Text, Runs, value, formatting);
            }

            private void AppendRevision(LegacyDocElement wrapper, LegacyDocRevisionKind kind) {
                LegacyDocRevision revision = ReadSupportedRevision(wrapper, kind);
                foreach (LegacyDocElement child in wrapper.Elements()) {
                    AppendChild(child, revision);
                }
            }
        }

        private readonly struct LegacyDocFieldBookmarkMarker {
            internal LegacyDocFieldBookmarkMarker(LegacyDocElement bookmarkStart, LegacyDocElement bookmarkEnd, int resultOffset) {
                BookmarkStart = bookmarkStart;
                BookmarkEnd = bookmarkEnd;
                ResultOffset = resultOffset;
            }

            internal LegacyDocElement BookmarkStart { get; }
            internal LegacyDocElement BookmarkEnd { get; }
            internal int ResultOffset { get; }
        }
    }
}
