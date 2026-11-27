using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using static OfficeIMO.Word.LegacyDoc.Docx.LegacyDocXml;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    internal sealed partial class LegacyDocDocxWriter {
        /// <summary>
        /// Where runs go: the paragraph they belong to, whose style decides relative toggles, and the element they are
        /// appended to, which is the paragraph itself or a hyperlink or field inside it.
        /// </summary>
        private readonly struct RunTarget {
            internal RunTarget(LegacyDocDocxPart part, int paragraph, int container, int imageRun = OfficeXmlTree.NoNode) {
                Part = part;
                Paragraph = paragraph;
                Container = container;
                ImageRun = imageRun;
            }

            internal LegacyDocDocxPart Part { get; }
            internal OfficeXmlTree Tree => Part.Tree;
            internal int Paragraph { get; }
            internal int Container { get; }

            /// <summary>
            /// The run pictures are added to, when the paragraph was created with a first text run: the engine's image
            /// API puts a picture into the paragraph's current run. Otherwise pictures get runs of their own.
            /// </summary>
            internal int ImageRun { get; }

            internal RunTarget In(int container) => new(Part, Paragraph, container, ImageRun);
        }

        private void AddRuns(RunTarget target, IReadOnlyList<LegacyDocTextRun> paragraphRuns, int startIndex, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            notes.CancellationToken.ThrowIfCancellationRequested();
            for (int index = startIndex; index < paragraphRuns.Count; index++) {
                notes.CancellationToken.ThrowIfCancellationRequested();
                LegacyDocTextRun legacyRun = paragraphRuns[index];
                if (legacyRun.Picture != null) {
                    AddPicture(target, legacyRun, bookmarks);
                    continue;
                }

                if (legacyRun.Revision.HasValue) {
                    AddRevisionRunContent(target, legacyRun, notes, bookmarks);
                    continue;
                }

                if (legacyRun.HyperlinkTarget.HasValue) {
                    int hyperlinkStartIndex = index;
                    LegacyDocHyperlinkTarget hyperlinkTarget = legacyRun.HyperlinkTarget;
                    while (index + 1 < paragraphRuns.Count && paragraphRuns[index + 1].HyperlinkTarget == hyperlinkTarget) {
                        notes.CancellationToken.ThrowIfCancellationRequested();
                        index++;
                    }

                    AddHyperlinkRunsContent(target, paragraphRuns, hyperlinkStartIndex, index - hyperlinkStartIndex + 1, notes, bookmarks);
                    continue;
                }

                if (legacyRun.IsPageNumber) {
                    AddPageNumber(target, legacyRun, bookmarks);
                    continue;
                }

                if (legacyRun.IsNumPages) {
                    AddPageCountField(target, legacyRun, bookmarks, " NUMPAGES  ");
                    continue;
                }

                if (legacyRun.IsSectionPages) {
                    AddPageCountField(target, legacyRun, bookmarks, string.IsNullOrWhiteSpace(legacyRun.FieldInstruction) ? " SECTIONPAGES  " : legacyRun.FieldInstruction!);
                    continue;
                }

                if (legacyRun.IsStaticDisplayField) {
                    AddStaticDisplayField(target, legacyRun, bookmarks);
                    continue;
                }

                AddRunContent(target, legacyRun, notes, bookmarks);
            }
        }

        private void AddRunContent(RunTarget target, LegacyDocTextRun legacyRun, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            notes.CancellationToken.ThrowIfCancellationRequested();
            if (legacyRun.IsPageNumber) {
                AddPageNumber(target, legacyRun, bookmarks);
                return;
            }

            if (legacyRun.IsNumPages) {
                AddPageCountField(target, legacyRun, bookmarks, " NUMPAGES  ");
                return;
            }

            if (legacyRun.IsSectionPages) {
                AddPageCountField(target, legacyRun, bookmarks, string.IsNullOrWhiteSpace(legacyRun.FieldInstruction) ? " SECTIONPAGES  " : legacyRun.FieldInstruction!);
                return;
            }

            if (legacyRun.IsStaticDisplayField) {
                AddStaticDisplayField(target, legacyRun, bookmarks);
                return;
            }

            if (legacyRun.HyperlinkTarget.HasValue) {
                AddHyperlinkRunsContent(target, new[] { legacyRun }, 0, 1, notes, bookmarks);
                return;
            }

            OfficeXmlTree tree = target.Tree;
            string text = legacyRun.Text;
            int segmentStart = 0;
            for (int index = 0; index < text.Length; index++) {
                notes.CancellationToken.ThrowIfCancellationRequested();
                char character = text[index];
                if (!IsSpecialRunCharacter(character)) {
                    int? markerPosition = RunCharacterPosition(legacyRun, index);
                    if (!bookmarks.HasMarkers(markerPosition)) {
                        continue;
                    }

                    AddTextSegment(target, legacyRun, segmentStart, index - segmentStart);
                    bookmarks.EmitAt(tree, target.Container, markerPosition);
                    segmentStart = index;
                    continue;
                }

                AddTextSegment(target, legacyRun, segmentStart, index - segmentStart);
                bookmarks.EmitAt(tree, target.Container, RunCharacterPosition(legacyRun, index));
                if (character == '\t') {
                    AddFormattedRun(target, legacyRun, Element(tree, "tab"));
                } else if (character == LegacyDocFootnoteReader.FootnoteReferenceCharacter) {
                    AddNoteReference(target, notes, RunCharacterPosition(legacyRun, index), legacyRun);
                } else if (character == LegacyDocCommentReader.CommentReferenceCharacter) {
                    AddCommentReference(target, notes, RunCharacterPosition(legacyRun, index), legacyRun);
                } else {
                    AddFormattedRun(target, legacyRun, CreateBreak(tree, BreakType(character)));
                }

                segmentStart = index + 1;
            }

            AddTextSegment(target, legacyRun, segmentStart, text.Length - segmentStart);
            bookmarks.EmitAt(tree, target.Container, RunEndCharacterPosition(legacyRun));
        }

        private static int? RunCharacterPosition(LegacyDocTextRun legacyRun, int index) =>
            index >= 0 && index < legacyRun.CharacterPositions.Count ? legacyRun.CharacterPositions[index] : null;

        private static int? RunEndCharacterPosition(LegacyDocTextRun legacyRun) =>
            legacyRun.CharacterPositions.Count == 0 ? null : legacyRun.CharacterPositions[legacyRun.CharacterPositions.Count - 1] + 1;

        private static bool ContainsSpecialRunCharacter(string text) {
            for (int index = 0; index < text.Length; index++) {
                if (IsSpecialRunCharacter(text[index])) return true;
            }

            return false;
        }

        private static bool IsSpecialRunCharacter(char character) =>
            character == '\t'
            || character == LegacyDocSpecialCharacters.TextWrappingBreak
            || character == LegacyDocSpecialCharacters.PageBreak
            || character == LegacyDocSpecialCharacters.ColumnBreak
            || character == LegacyDocFootnoteReader.FootnoteReferenceCharacter
            || character == LegacyDocCommentReader.CommentReferenceCharacter;

        private static bool IsHyperlinkSpecialRunCharacter(char character) =>
            character == '\t'
            || character == LegacyDocSpecialCharacters.TextWrappingBreak
            || character == LegacyDocSpecialCharacters.PageBreak
            || character == LegacyDocSpecialCharacters.ColumnBreak;

        private static bool IsFieldResultSpecialRunCharacter(char character) =>
            character == '\t'
            || character == LegacyDocSpecialCharacters.SoftHyphen
            || character == LegacyDocSpecialCharacters.NoBreakHyphen
            || character == LegacyDocSpecialCharacters.TextWrappingBreak
            || character == LegacyDocSpecialCharacters.PageBreak
            || character == LegacyDocSpecialCharacters.ColumnBreak;

        private static string? BreakType(char character) {
            if (character == LegacyDocSpecialCharacters.PageBreak) return "page";
            if (character == LegacyDocSpecialCharacters.ColumnBreak) return "column";
            return null;
        }

        private static int CreateBreak(OfficeXmlTree tree, string? type) {
            int element = Element(tree, "br");
            Set(tree, element, "type", type);
            return element;
        }

        /// <summary>Creates a run holding <paramref name="content"/>.</summary>
        private static int CreateRun(OfficeXmlTree tree, int content) {
            int run = Element(tree, "r");
            Append(tree, run, content);
            return run;
        }

        /// <summary>Appends a run holding <paramref name="content"/> to the container and formats it.</summary>
        private int AddFormattedRun(RunTarget target, LegacyDocTextRun legacyRun, int content) {
            int run = Append(target.Tree, target.Container, CreateRun(target.Tree, content));
            ApplyRunFormatting(target, run, legacyRun);
            return run;
        }

        private void AddTextSegment(RunTarget target, LegacyDocTextRun legacyRun, int startIndex, int length) {
            if (length <= 0) return;
            AddFormattedRun(target, legacyRun, TextElement(target.Tree, legacyRun.Text.Substring(startIndex, length)));
        }

        // ---- Hyperlinks -----------------------------------------------------------------------------------------

        private void AddHyperlinkRunsContent(RunTarget target, IReadOnlyList<LegacyDocTextRun> legacyRuns, int startIndex, int count, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            LegacyDocTextRun firstRun = legacyRuns[startIndex];
            LegacyDocHyperlinkTarget hyperlinkTarget = firstRun.HyperlinkTarget;
            OfficeXmlTree tree = target.Tree;
            int hyperlink;
            if (hyperlinkTarget.Uri != null && Uri.TryCreate(hyperlinkTarget.Uri, UriKind.Absolute, out Uri? uri)) {
                string id = _package.AddRelationship(target.Part.Name, RelationshipTypePrefix + "hyperlink", uri.OriginalString, external: true, id: NextRelationshipId());
                hyperlink = Element(tree, "hyperlink");
                tree.SetAttribute(hyperlink, OfficeXmlNamespaces.Relationships, "id", id);
                Set(tree, hyperlink, "history", "true");
            } else if (hyperlinkTarget.Anchor != null) {
                hyperlink = Element(tree, "hyperlink");
                Set(tree, hyperlink, "anchor", hyperlinkTarget.Anchor);
                Set(tree, hyperlink, "history", "true");
            } else {
                for (int index = startIndex; index < startIndex + count; index++) {
                    AddRunContent(target, CreateRunWithoutHyperlink(legacyRuns[index]), notes, bookmarks);
                }

                return;
            }

            Set(tree, hyperlink, "tooltip", hyperlinkTarget.Tooltip);
            Set(tree, hyperlink, "tgtFrame", hyperlinkTarget.TargetFrame);
            RunTarget inner = target.In(hyperlink);
            for (int index = startIndex; index < startIndex + count; index++) {
                AppendHyperlinkRunContent(inner, legacyRuns[index], bookmarks);
            }

            Append(tree, target.Container, hyperlink);
        }

        private void AppendHyperlinkRunContent(RunTarget target, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            OfficeXmlTree tree = target.Tree;
            if (legacyRun.IsStaticDisplayField) {
                bookmarks.EmitAt(tree, target.Container, RunCharacterPosition(legacyRun, 0));
                int simpleField = CreateSimpleField(tree, string.IsNullOrWhiteSpace(legacyRun.FieldInstruction)
                    ? StaticFieldInstruction(legacyRun.FieldKind)
                    : legacyRun.FieldInstruction!);
                AppendFieldResultContent(target, simpleField, legacyRun, legacyRun.Text, bookmarks);
                Append(tree, target.Container, simpleField);
                bookmarks.EmitAt(tree, target.Container, RunEndCharacterPosition(legacyRun));
                return;
            }

            string text = legacyRun.Text;
            int segmentStart = 0;
            for (int index = 0; index < text.Length; index++) {
                char character = text[index];
                if (!IsHyperlinkSpecialRunCharacter(character)) {
                    int? markerPosition = RunCharacterPosition(legacyRun, index);
                    if (!bookmarks.HasMarkers(markerPosition)) {
                        continue;
                    }

                    AddTextSegment(target, legacyRun, segmentStart, index - segmentStart);
                    bookmarks.EmitAt(tree, target.Container, markerPosition);
                    segmentStart = index;
                    continue;
                }

                AddTextSegment(target, legacyRun, segmentStart, index - segmentStart);
                bookmarks.EmitAt(tree, target.Container, RunCharacterPosition(legacyRun, index));
                AddFormattedRun(target, legacyRun, character == '\t' ? Element(tree, "tab") : CreateBreak(tree, BreakType(character)));
                segmentStart = index + 1;
            }

            AddTextSegment(target, legacyRun, segmentStart, text.Length - segmentStart);
            bookmarks.EmitAt(tree, target.Container, RunEndCharacterPosition(legacyRun));
        }

        private static LegacyDocTextRun CreateRunWithoutHyperlink(LegacyDocTextRun source) {
            return new LegacyDocTextRun(
                source.Text,
                source.Bold,
                source.Italic,
                source.Strike,
                source.DoubleStrike,
                source.Outline,
                source.Shadow,
                source.Emboss,
                source.Imprint,
                source.Hidden,
                source.NoProof,
                source.Caps,
                source.VerticalPosition,
                source.Underline,
                source.Highlight,
                source.FontSizeHalfPoints,
                source.ColorHex,
                source.FontFamily,
                source.CharacterPositions,
                fieldKind: source.FieldKind,
                fieldInstruction: source.FieldInstruction,
                fieldResultRuns: source.FieldResultRuns,
                fieldLocked: source.FieldLocked,
                specified: source.Specified,
                styleRelative: source.StyleRelative,
                styleInverted: source.StyleInverted,
                characterSpacingTwips: source.CharacterSpacingTwips,
                characterScalePercentage: source.CharacterScalePercentage,
                kerningMinimumFontSizeHalfPoints: source.KerningMinimumFontSizeHalfPoints,
                language: source.Language,
                eastAsiaLanguage: source.EastAsiaLanguage,
                picture: source.Picture,
                revision: source.Revision);
        }

        // ---- Fields ---------------------------------------------------------------------------------------------

        private static int CreateSimpleField(OfficeXmlTree tree, string instruction) {
            int field = Element(tree, "fldSimple");
            Set(tree, field, "instr", instruction);
            return field;
        }

        private void AddPageNumber(RunTarget target, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            if (legacyRun.FieldResultRuns.Count > 0 || legacyRun.FieldLocked == true) {
                AddPageCountField(target, legacyRun, bookmarks, " PAGE ");
                return;
            }

            OfficeXmlTree tree = target.Tree;
            bookmarks.EmitAt(tree, target.Container, RunCharacterPosition(legacyRun, 0));
            AddFormattedRun(target, legacyRun, Element(tree, "pgNum"));
            bookmarks.EmitAt(tree, target.Container, RunEndCharacterPosition(legacyRun));
        }

        // Import keeps the producer's cached count; the field stays available for later layout or evaluation.
        private void AddPageCountField(RunTarget target, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks, string instruction) {
            OfficeXmlTree tree = target.Tree;
            bookmarks.EmitAt(tree, target.Container, RunCharacterPosition(legacyRun, 0));
            int simpleField = CreateSimpleField(tree, legacyRun.FieldInstruction ?? instruction);
            AppendFieldResultContent(target, simpleField, legacyRun, string.IsNullOrEmpty(legacyRun.Text) ? "1" : legacyRun.Text, bookmarks);
            Append(tree, target.Container, simpleField);
            bookmarks.EmitAt(tree, target.Container, RunEndCharacterPosition(legacyRun));
        }

        private void AddStaticDisplayField(RunTarget target, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            OfficeXmlTree tree = target.Tree;
            bookmarks.EmitAt(tree, target.Container, RunCharacterPosition(legacyRun, 0));
            int simpleField = CreateSimpleField(tree, string.IsNullOrWhiteSpace(legacyRun.FieldInstruction)
                ? StaticFieldInstruction(legacyRun.FieldKind)
                : legacyRun.FieldInstruction!);
            AppendFieldResultContent(target, simpleField, legacyRun, legacyRun.Text, bookmarks);
            Append(tree, target.Container, simpleField);
            bookmarks.EmitAt(tree, target.Container, RunEndCharacterPosition(legacyRun));
        }

        private void AppendFieldResultContent(RunTarget target, int simpleField, LegacyDocTextRun legacyRun, string resultText, LegacyDocBookmarkProjection bookmarks) {
            OfficeXmlTree tree = target.Tree;
            if (legacyRun.FieldLocked == true) Set(tree, simpleField, "fldLock", "true");
            RunTarget inner = target.In(simpleField);
            if (legacyRun.FieldResultRuns.Count > 0) {
                AddRuns(inner, legacyRun.FieldResultRuns, 0, LegacyDocNoteProjection.Empty, bookmarks);
                NormalizeFieldResultHyphens(tree, simpleField);
                return;
            }

            int segmentStart = 0;
            for (int index = 0; index < resultText.Length; index++) {
                char character = resultText[index];
                if (!IsFieldResultSpecialRunCharacter(character)) continue;
                AppendFieldTextSegment(inner, legacyRun, resultText, segmentStart, index - segmentStart);
                AddFormattedRun(inner, legacyRun, CreateFieldSpecialContent(tree, character));
                segmentStart = index + 1;
            }

            AppendFieldTextSegment(inner, legacyRun, resultText, segmentStart, resultText.Length - segmentStart);
            if (!HasElements(tree, simpleField)) {
                AppendFieldTextSegment(inner, legacyRun, "1", 0, 1);
            }
        }

        private void AppendFieldTextSegment(RunTarget target, LegacyDocTextRun legacyRun, string text, int startIndex, int length) {
            if (length <= 0) return;
            AddFormattedRun(target, legacyRun, TextElement(target.Tree, text.Substring(startIndex, length)));
        }

        private static int CreateFieldSpecialContent(OfficeXmlTree tree, char character) {
            if (character == '\t') return Element(tree, "tab");
            if (character == LegacyDocSpecialCharacters.SoftHyphen) return Element(tree, "softHyphen");
            if (character == LegacyDocSpecialCharacters.NoBreakHyphen) return Element(tree, "noBreakHyphen");
            return CreateBreak(tree, character == LegacyDocSpecialCharacters.TextWrappingBreak ? null : BreakType(character));
        }

        // The shared run writer keeps formatting and revisions but writes literal hyphens in ordinary text. Field
        // caches keep the native special-character elements.
        private static void NormalizeFieldResultHyphens(OfficeXmlTree tree, int field) {
            foreach (int text in Descendants(tree, field)) {
                if (!Is(tree, text, "t") && !Is(tree, text, "delText")) continue;
                string value = tree.GetInnerText(text);
                int parent = tree.ParentOf(text);
                int segmentStart = 0;
                for (int index = 0; index < value.Length; index++) {
                    char character = value[index];
                    if (character != LegacyDocSpecialCharacters.SoftHyphen && character != LegacyDocSpecialCharacters.NoBreakHyphen) continue;
                    if (index > segmentStart) InsertText(segmentStart, index - segmentStart);
                    tree.InsertBefore(parent, CreateFieldSpecialContent(tree, character), text);
                    segmentStart = index + 1;
                }

                if (segmentStart == 0) continue;
                if (segmentStart < value.Length) InsertText(segmentStart, value.Length - segmentStart);
                tree.Remove(text);

                void InsertText(int start, int length) {
                    int segment = TextElement(tree, value.Substring(start, length), tree.LocalNameOf(text));
                    tree.InsertBefore(parent, segment, text);
                }
            }
        }

        private static string StaticFieldInstruction(LegacyDocFieldKind fieldKind) {
            return fieldKind switch {
                LegacyDocFieldKind.Time => " TIME  ",
                LegacyDocFieldKind.CreateDate => " CREATEDATE  ",
                LegacyDocFieldKind.SaveDate => " SAVEDATE  ",
                LegacyDocFieldKind.PrintDate => " PRINTDATE  ",
                LegacyDocFieldKind.DocumentProperty => throw new NotSupportedException("Legacy DOC document-property field projection requires the source field instruction."),
                LegacyDocFieldKind.TextReference => throw new NotSupportedException("Legacy DOC text-reference field projection requires the source field instruction."),
                LegacyDocFieldKind.Equation => " EQ ",
                _ => " DATE  "
            };
        }

        // ---- Revisions ------------------------------------------------------------------------------------------

        private void AddRevisionRunContent(RunTarget target, LegacyDocTextRun legacyRun, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            OfficeXmlTree tree = target.Tree;
            string text = legacyRun.Text;
            int segmentStart = 0;
            for (int index = 0; index < text.Length; index++) {
                char character = text[index];
                int? markerPosition = RunCharacterPosition(legacyRun, index);
                if (!IsSpecialRunCharacter(character)) {
                    if (!bookmarks.HasMarkers(markerPosition)) continue;
                    AddRevisionTextSegment(target, legacyRun, segmentStart, index - segmentStart);
                    bookmarks.EmitAt(tree, target.Container, markerPosition);
                    segmentStart = index;
                    continue;
                }

                AddRevisionTextSegment(target, legacyRun, segmentStart, index - segmentStart);
                bookmarks.EmitAt(tree, target.Container, markerPosition);
                if (character == '\t') {
                    AddRevisionElement(target, legacyRun, Element(tree, "tab"));
                } else if (character == LegacyDocFootnoteReader.FootnoteReferenceCharacter) {
                    AddNoteReference(target, notes, markerPosition, legacyRun);
                } else if (character == LegacyDocCommentReader.CommentReferenceCharacter) {
                    AddCommentReference(target, notes, markerPosition, legacyRun);
                } else {
                    AddRevisionElement(target, legacyRun, CreateBreak(tree, BreakType(character)));
                }

                segmentStart = index + 1;
            }

            AddRevisionTextSegment(target, legacyRun, segmentStart, text.Length - segmentStart);
            bookmarks.EmitAt(tree, target.Container, RunEndCharacterPosition(legacyRun));
        }

        private void AddRevisionTextSegment(RunTarget target, LegacyDocTextRun legacyRun, int startIndex, int length) {
            if (length <= 0) return;
            string value = legacyRun.Text.Substring(startIndex, length);
            AddRevisionElement(target, legacyRun, TextElement(target.Tree, value, legacyRun.Revision.Kind == LegacyDocRevisionKind.Deleted ? "delText" : "t"));
        }

        private void AddRevisionElement(RunTarget target, LegacyDocTextRun legacyRun, int content) {
            int run = CreateRun(target.Tree, content);
            ApplyRunFormatting(target, run, legacyRun);
            AppendRevisionRun(target, legacyRun, run);
        }

        /// <summary>Formats a note or comment reference run, and moves it into a revision when the source run has one.</summary>
        private void ApplyReferenceFormatting(RunTarget target, int run, LegacyDocTextRun legacyRun) {
            ApplyRunFormatting(target, run, legacyRun);
            if (legacyRun.Revision.HasValue) {
                target.Tree.Remove(run);
                AppendRevisionRun(target, legacyRun, run);
            }
        }

        private void AppendRevisionRun(RunTarget target, LegacyDocTextRun legacyRun, int run) {
            OfficeXmlTree tree = target.Tree;
            int revision = Element(tree, legacyRun.Revision.Kind == LegacyDocRevisionKind.Deleted ? "del" : "ins");
            Set(tree, revision, "id", WordHeadersAndFooters.GenerateRevisionId());
            Set(tree, revision, "author", legacyRun.Revision.Author ?? LegacyDocRevisionAuthorReader.UnknownAuthor);
            if (legacyRun.Revision.Date != null) {
                Set(tree, revision, "date", System.Xml.XmlConvert.ToString(legacyRun.Revision.Date.Value, System.Xml.XmlDateTimeSerializationMode.RoundtripKind));
            }

            Append(tree, revision, run);
            Append(tree, target.Container, revision);
        }
    }
}
