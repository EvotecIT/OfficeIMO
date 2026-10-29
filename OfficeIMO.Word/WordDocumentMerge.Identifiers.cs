using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal sealed partial class WordDocumentMerge {
        private const int MaxBookmarkNameLength = 40;
        private static readonly Regex FieldInstructionToken = new("\"[^\"]*\"|[^\\s\"]+", RegexOptions.CultureInvariant);

        /// <summary>
        /// Gives imported bookmarks new ids and renames bookmarks whose names already exist in the destination,
        /// updating hyperlinks and field instructions that target them.
        /// </summary>
        private static void RemapBookmarks(IReadOnlyList<OpenXmlElement> roots, DestinationIdentifiers identifiers) {
            var idMap = new Dictionary<string, string>(StringComparer.Ordinal);
            var nameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            string MapId(string oldId) {
                if (!idMap.TryGetValue(oldId, out string? newId)) {
                    newId = identifiers.CreateBookmarkId();
                    idMap[oldId] = newId;
                }

                return newId;
            }

            foreach (OpenXmlElement root in roots) {
                foreach (BookmarkStart start in root.Descendants<BookmarkStart>()) {
                    if (start.Id?.Value is string oldId) {
                        start.Id = MapId(oldId);
                    }

                    string? name = start.Name?.Value;
                    if (string.IsNullOrEmpty(name) || identifiers.BookmarkNames.Add(name!)) {
                        continue;
                    }

                    string renamed = identifiers.CreateBookmarkName(name!);
                    nameMap[name!] = renamed;
                    start.Name = renamed;
                }
            }

            foreach (OpenXmlElement root in roots) {
                foreach (BookmarkEnd end in root.Descendants<BookmarkEnd>()) {
                    if (end.Id?.Value is string oldId) {
                        end.Id = MapId(oldId);
                    }
                }
            }

            if (nameMap.Count == 0) {
                return;
            }

            foreach (OpenXmlElement root in roots) {
                foreach (Hyperlink hyperlink in root.Descendants<Hyperlink>()) {
                    if (hyperlink.Anchor?.Value is string anchor && nameMap.TryGetValue(anchor, out string? renamed)) {
                        hyperlink.Anchor = renamed;
                    }
                }

                foreach (SimpleField field in root.Descendants<SimpleField>()) {
                    if (field.Instruction?.Value is string instruction) {
                        field.Instruction = RenameBookmarkTokens(instruction, nameMap);
                    }
                }

                RewriteComplexBookmarkInstructions(root, nameMap);
            }
        }

        private sealed class FieldInstructionGroup {
            internal List<FieldCode?> Codes { get; } = new();
            internal bool InInstruction { get; set; } = true;
        }

        private static void RewriteComplexBookmarkInstructions(OpenXmlElement root, Dictionary<string, string> nameMap) {
            var fields = new Stack<FieldInstructionGroup>();
            foreach (OpenXmlElement element in root.Descendants()) {
                if (element is FieldChar marker) {
                    if (marker.FieldCharType?.Value == FieldCharValues.Begin) {
                        if (fields.Count > 0 && fields.Peek().InInstruction) fields.Peek().Codes.Add(null);
                        fields.Push(new FieldInstructionGroup());
                    } else if (marker.FieldCharType?.Value == FieldCharValues.Separate && fields.Count > 0) {
                        RewriteFieldCodeRuns(fields.Peek(), nameMap);
                        fields.Peek().InInstruction = false;
                    } else if (marker.FieldCharType?.Value == FieldCharValues.End && fields.Count > 0) {
                        RewriteFieldCodeRuns(fields.Pop(), nameMap);
                    }
                } else if (element is FieldCode code) {
                    if (fields.Count > 0 && fields.Peek().InInstruction) fields.Peek().Codes.Add(code);
                    else if (fields.Count == 0) code.Text = RenameBookmarkTokens(code.Text, nameMap);
                }
            }
        }

        private static void RewriteFieldCodeRuns(FieldInstructionGroup field, Dictionary<string, string> nameMap) {
            if (!field.InInstruction || field.Codes.Count == 0) return;
            const string nested = " \u0001 ";
            string original = string.Concat(field.Codes.Select(x => x?.Text ?? nested));
            string rewritten = RenameBookmarkTokens(original, nameMap);
            if (string.Equals(original, rewritten, StringComparison.Ordinal)) return;
            int start = 0;
            while (start < original.Length && start < rewritten.Length && original[start] == rewritten[start]) start++;
            int suffix = 0;
            while (suffix < original.Length - start && suffix < rewritten.Length - start
                && original[original.Length - suffix - 1] == rewritten[rewritten.Length - suffix - 1]) suffix++;
            int end = original.Length - suffix;
            string replacement = rewritten.Substring(start, rewritten.Length - start - suffix);
            int offset = 0;
            bool inserted = false;
            foreach (FieldCode? code in field.Codes) {
                string text = code?.Text ?? nested;
                int next = offset + text.Length;
                if (code != null && ((end == start && !inserted && offset <= start && start <= next)
                    || (next > start && offset < end))) {
                    int localStart = Math.Max(start - offset, 0);
                    int localEnd = Math.Min(end - offset, text.Length);
                    code.Text = text.Substring(0, localStart) + (inserted ? string.Empty : replacement) + text.Substring(localEnd);
                    inserted = true;
                }
                offset = next;
            }
        }

        private static string RenameBookmarkTokens(string instruction, Dictionary<string, string> nameMap) {
            MatchCollection tokens = FieldInstructionToken.Matches(instruction);
            if (tokens.Count < 2) return instruction;
            string kind = tokens[0].Value;
            int operand = -1;
            if (string.Equals(kind, "REF", StringComparison.OrdinalIgnoreCase)
                || string.Equals(kind, "PAGEREF", StringComparison.OrdinalIgnoreCase)
                || string.Equals(kind, "NOTEREF", StringComparison.OrdinalIgnoreCase)) {
                operand = 1;
            } else if (string.Equals(kind, "HYPERLINK", StringComparison.OrdinalIgnoreCase)) {
                for (int i = 1; i + 1 < tokens.Count; i++) {
                    if (string.Equals(tokens[i].Value, @"\l", StringComparison.OrdinalIgnoreCase)) { operand = i + 1; break; }
                }
            }
            if (operand < 0) return instruction;
            Match match = tokens[operand];
            string token = match.Value;
            bool quoted = token.Length >= 2 && token[0] == '"' && token[token.Length - 1] == '"';
            string bare = quoted ? token.Substring(1, token.Length - 2) : token;
            if (!nameMap.TryGetValue(bare, out string? renamed)) return instruction;
            string replacement = quoted ? "\"" + renamed + "\"" : renamed;
            return instruction.Substring(0, match.Index) + replacement + instruction.Substring(match.Index + match.Length);
        }

        /// <summary>
        /// Gives imported paragraphs and table rows new <c>w14:paraId</c> values where the destination already uses them.
        /// </summary>
        /// <returns>The replaced paragraph ids, keyed by their source value.</returns>
        private static Dictionary<string, string> RemapParagraphIds(IReadOnlyList<OpenXmlElement> roots, DestinationIdentifiers identifiers) {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (OpenXmlElement root in roots) {
                foreach (OpenXmlElement element in root.Descendants()) {
                    string? value = element switch {
                        Paragraph paragraph => paragraph.ParagraphId?.Value,
                        TableRow row => row.ParagraphId?.Value,
                        _ => null
                    };
                    if (string.IsNullOrEmpty(value) || identifiers.ParagraphIds.Add(value!)) {
                        continue;
                    }

                    string replacement = identifiers.CreateParagraphId();
                    if (element is Paragraph paragraphElement) {
                        paragraphElement.ParagraphId = replacement;
                    } else if (element is TableRow rowElement) {
                        rowElement.ParagraphId = replacement;
                    }

                    map[value!] = replacement;
                }
            }

            return map;
        }

        /// <summary>
        /// Identifiers in use by the destination before the merge.
        /// </summary>
        private sealed class DestinationIdentifiers {
            private int _nextBookmarkId;
            private uint _nextParagraphId;

            internal DestinationIdentifiers(MainDocumentPart main) {
                var roots = new List<OpenXmlElement?> {
                    main.Document,
                    main.FootnotesPart?.Footnotes,
                    main.EndnotesPart?.Endnotes,
                    main.WordprocessingCommentsPart?.Comments
                };
                roots.AddRange(main.HeaderParts.Select(_ => (OpenXmlElement?)_.Header));
                roots.AddRange(main.FooterParts.Select(_ => (OpenXmlElement?)_.Footer));

                int maxBookmarkId = -1;
                foreach (OpenXmlElement root in roots.OfType<OpenXmlElement>()) {
                    foreach (OpenXmlElement element in root.Descendants()) {
                        switch (element) {
                            case BookmarkStart start:
                                if (start.Name?.Value is string name) {
                                    BookmarkNames.Add(name);
                                }

                                maxBookmarkId = Math.Max(maxBookmarkId, ParseId(start.Id?.Value));
                                break;
                            case BookmarkEnd end:
                                maxBookmarkId = Math.Max(maxBookmarkId, ParseId(end.Id?.Value));
                                break;
                            case Paragraph paragraph when paragraph.ParagraphId?.Value is string paragraphId:
                                ParagraphIds.Add(paragraphId);
                                break;
                            case TableRow row when row.ParagraphId?.Value is string rowId:
                                ParagraphIds.Add(rowId);
                                break;
                        }
                    }
                }

                _nextBookmarkId = maxBookmarkId + 1;
            }

            internal HashSet<string> BookmarkNames { get; } = new(StringComparer.OrdinalIgnoreCase);

            internal HashSet<string> ParagraphIds { get; } = new(StringComparer.OrdinalIgnoreCase);

            internal string CreateBookmarkId() {
                return (_nextBookmarkId++).ToString(CultureInfo.InvariantCulture);
            }

            internal string CreateBookmarkName(string name) {
                for (int suffix = 1; ; suffix++) {
                    string tail = "_" + suffix.ToString(CultureInfo.InvariantCulture);
                    string stem = name.Length + tail.Length > MaxBookmarkNameLength
                        ? name.Substring(0, Math.Max(0, MaxBookmarkNameLength - tail.Length))
                        : name;
                    string candidate = stem + tail;
                    if (BookmarkNames.Add(candidate)) {
                        return candidate;
                    }
                }
            }

            internal string CreateParagraphId() {
                // Paragraph ids must stay below 0x80000000.
                string candidate;
                do {
                    _nextParagraphId = _nextParagraphId >= 0x7FFFFFFFU ? 1U : _nextParagraphId + 1U;
                    candidate = _nextParagraphId.ToString("X8", CultureInfo.InvariantCulture);
                } while (!ParagraphIds.Add(candidate));

                return candidate;
            }

            private static int ParseId(string? value) {
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)) {
                    return parsed;
                }

                return -1;
            }
        }
    }
}
