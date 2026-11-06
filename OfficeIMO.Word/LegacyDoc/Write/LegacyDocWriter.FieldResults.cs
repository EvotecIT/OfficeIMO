using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
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
                if (marker.BookmarkStart != null) {
                    bookmarks.AddStart(marker.BookmarkStart, characterPosition);
                } else if (marker.BookmarkEnd != null) {
                    bookmarks.AddEnd(marker.BookmarkEnd, characterPosition);
                }
            }
            int endCharacter = characterOffset + text.Length;
            text.Append(LegacyDocField.End);
            runs.Add(new LegacyDocWritableRun(endCharacter, 1, LegacyDocWritableFormatting.SpecialCharacter, fieldLocked: result.FieldLocked));
        }

        private static LegacyDocWritableFieldResult ReadSimpleFieldResult(
            SimpleField field,
            bool allowHyperlinkRunStyle = false,
            LegacyDocWritableFormatting inheritedFormatting = default) {
            var result = new LegacyDocWritableFieldResult(inheritedFormatting, allowHyperlinkRunStyle) { FieldLocked = field.FieldLock?.Value == true };
            foreach (OpenXmlElement child in field.ChildElements) {
                result.AppendChild(child);
            }
            return result;
        }

        private static void AppendSupportedFieldResultElementText(StringBuilder result, OpenXmlElement element) {
            switch (element) {
                case Text text:
                    result.Append(text.Text);
                    break;
                case DeletedText deletedText:
                    result.Append(deletedText.Text);
                    break;
                case TabChar:
                    result.Append('\t');
                    break;
                case CarriageReturn:
                    result.Append(LegacyDocSpecialCharacters.TextWrappingBreak);
                    break;
                case NoBreakHyphen:
                    result.Append(LegacyDocSpecialCharacters.NoBreakHyphen);
                    break;
                case SoftHyphen:
                    result.Append(LegacyDocSpecialCharacters.SoftHyphen);
                    break;
                case Break breakNode:
                    result.Append(GetSupportedFieldResultBreakCharacter(breakNode));
                    break;
                default:
                    throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} field result runs only when they contain text, tabs, carriage returns, soft/no-break hyphens, and supported breaks. Unsupported field result element: {element.LocalName}.");
            }
        }

        private static char GetSupportedFieldResultBreakCharacter(Break breakNode) {
            BreakValues? breakType = breakNode.Type?.Value;
            if (breakType == null || breakType == BreakValues.TextWrapping) {
                return LegacyDocSpecialCharacters.TextWrappingBreak;
            }
            if (breakType == BreakValues.Page) {
                return LegacyDocSpecialCharacters.PageBreak;
            }
            if (breakType == BreakValues.Column) {
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

            internal void AppendChild(OpenXmlElement child, LegacyDocRevision revision = default) {
                switch (child) {
                    case Run run:
                        LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties, _allowHyperlinkRunStyle)
                            .WithInheritedFormatting(_inheritedFormatting);
                        if (revision.HasValue) formatting = formatting.WithRevision(revision);
                        _emptyFormatting ??= formatting;
                        var runText = new StringBuilder();
                        foreach (OpenXmlElement element in run.ChildElements) {
                            if (element is RunProperties || element is LastRenderedPageBreak) continue;
                            AppendSupportedFieldResultElementText(runText, element);
                        }
                        AppendText(runText.ToString(), formatting);
                        break;
                    case BookmarkStart bookmarkStart:
                        BookmarkMarkers.Add(new LegacyDocFieldBookmarkMarker(bookmarkStart, null, Text.Length));
                        break;
                    case BookmarkEnd bookmarkEnd:
                        BookmarkMarkers.Add(new LegacyDocFieldBookmarkMarker(null, bookmarkEnd, Text.Length));
                        break;
                    case InsertedRun inserted when !revision.HasValue && !_inheritedFormatting.Revision.HasValue:
                        AppendRevision(inserted, LegacyDocRevisionKind.Inserted);
                        break;
                    case DeletedRun deleted when !revision.HasValue && !_inheritedFormatting.Revision.HasValue:
                        AppendRevision(deleted, LegacyDocRevisionKind.Deleted);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) return;
                        throw new NotSupportedException($"Native DOC saving supports {SupportedFieldNames} field results only with text runs, bookmarks, and basic tracked insertions or deletions. Nested fields and nested revisions are not supported. Unsupported field result element: {child.LocalName}.");
                }
            }

            internal void AppendElement(OpenXmlElement element, LegacyDocWritableFormatting formatting) {
                var value = new StringBuilder();
                AppendSupportedFieldResultElementText(value, element);
                AppendText(value.ToString(), formatting);
            }

            internal void AppendText(string value, LegacyDocWritableFormatting formatting) {
                _emptyFormatting ??= formatting;
                AppendFormattedText(Text, Runs, value, formatting);
            }

            private void AppendRevision(OpenXmlCompositeElement wrapper, LegacyDocRevisionKind kind) {
                LegacyDocRevision revision = ReadSupportedRevision(wrapper, kind);
                foreach (OpenXmlElement child in wrapper.ChildElements) {
                    AppendChild(child, revision);
                }
            }
        }

        private readonly struct LegacyDocFieldBookmarkMarker {
            internal LegacyDocFieldBookmarkMarker(BookmarkStart? bookmarkStart, BookmarkEnd? bookmarkEnd, int resultOffset) {
                BookmarkStart = bookmarkStart;
                BookmarkEnd = bookmarkEnd;
                ResultOffset = resultOffset;
            }

            internal BookmarkStart? BookmarkStart { get; }
            internal BookmarkEnd? BookmarkEnd { get; }
            internal int ResultOffset { get; }
        }
    }
}
