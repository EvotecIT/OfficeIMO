using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static bool IsLegacyDocFieldResultSpecialRunCharacter(char character) {
            return character == '\t'
                || character == LegacyDocSpecialCharacters.SoftHyphen
                || character == LegacyDocSpecialCharacters.NoBreakHyphen
                || character == LegacyDocSpecialCharacters.TextWrappingBreak
                || character == LegacyDocSpecialCharacters.PageBreak
                || character == LegacyDocSpecialCharacters.ColumnBreak;
        }

        private static void AddLegacyDocStaticDisplayField(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunCharacterPosition(legacyRun, 0));
            var simpleField = new SimpleField { Instruction = string.IsNullOrWhiteSpace(legacyRun.FieldInstruction) ? GetLegacyDocStaticFieldInstruction(legacyRun.FieldKind) : legacyRun.FieldInstruction };
            AppendLegacyDocFieldResultContent(simpleField, paragraph, legacyRun, legacyRun.Text, bookmarks);
            paragraph._paragraph.Append(simpleField);
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunEndCharacterPosition(legacyRun));
        }

        private static void AppendLegacyDocFieldResultContent(SimpleField simpleField, WordParagraph paragraph, LegacyDocTextRun legacyRun, string resultText, LegacyDocBookmarkProjection bookmarks) {
            if (legacyRun.FieldLocked == true) simpleField.FieldLock = true;
            if (legacyRun.FieldResultRuns.Count > 0) {
                int originalChildren = paragraph._paragraph.ChildElements.Count;
                AddLegacyDocRuns(paragraph, legacyRun.FieldResultRuns, LegacyDocNoteProjection.Empty, bookmarks);
                foreach (OpenXmlElement child in paragraph._paragraph.ChildElements.Skip(originalChildren).ToArray()) {
                    child.Remove();
                    simpleField.Append(child);
                }
                NormalizeLegacyDocFieldResultHyphens(simpleField);
                return;
            }
            int segmentStart = 0;
            for (int index = 0; index < resultText.Length; index++) {
                char character = resultText[index];
                if (!IsLegacyDocFieldResultSpecialRunCharacter(character)) {
                    continue;
                }

                AppendLegacyDocFieldTextSegment(simpleField, paragraph, legacyRun, resultText, segmentStart, index - segmentStart);
                AppendLegacyDocFieldSpecialRun(simpleField, paragraph, legacyRun, character);
                segmentStart = index + 1;
            }

            AppendLegacyDocFieldTextSegment(simpleField, paragraph, legacyRun, resultText, segmentStart, resultText.Length - segmentStart);
            if (!simpleField.HasChildren) {
                AppendLegacyDocFieldTextSegment(simpleField, paragraph, legacyRun, "1", 0, 1);
            }
        }

        private static void AppendLegacyDocFieldTextSegment(SimpleField simpleField, WordParagraph paragraph, LegacyDocTextRun legacyRun, string text, int startIndex, int length) {
            if (length <= 0) {
                return;
            }

            var run = new Run(new Text(text.Substring(startIndex, length)) {
                Space = SpaceProcessingModeValues.Preserve
            });
            simpleField.Append(run);
            ApplyLegacyDocRunFormatting(new WordParagraph(paragraph._document, paragraph._paragraph, run), legacyRun);
        }

        private static void AppendLegacyDocFieldSpecialRun(SimpleField simpleField, WordParagraph paragraph, LegacyDocTextRun legacyRun, char character) {
            var run = new Run(CreateLegacyDocFieldSpecialContent(character));
            simpleField.Append(run);
            ApplyLegacyDocRunFormatting(new WordParagraph(paragraph._document, paragraph._paragraph, run), legacyRun);
        }

        private static OpenXmlElement CreateLegacyDocFieldSpecialContent(char character) {
            if (character == '\t') return new TabChar();
            if (character == LegacyDocSpecialCharacters.SoftHyphen) return new SoftHyphen();
            if (character == LegacyDocSpecialCharacters.NoBreakHyphen) return new NoBreakHyphen();
            return character == LegacyDocSpecialCharacters.TextWrappingBreak
                ? new Break() : new Break { Type = GetLegacyDocBreakType(character) };
        }

        private static void NormalizeLegacyDocFieldResultHyphens(SimpleField field) {
            // The shared run projector retains formatting/revisions but ordinary text uses
            // literal hyphens. Field caches retain the native special-character elements.
            foreach (TextType text in field.Descendants<TextType>()
                .Where(text => text is Text or DeletedText).ToArray()) {
                string value = text.Text;
                int segmentStart = 0;
                for (int index = 0; index < value.Length; index++) {
                    char character = value[index];
                    if (character != LegacyDocSpecialCharacters.SoftHyphen
                        && character != LegacyDocSpecialCharacters.NoBreakHyphen) continue;
                    if (index > segmentStart) AppendText(segmentStart, index - segmentStart);
                    text.InsertBeforeSelf(CreateLegacyDocFieldSpecialContent(character));
                    segmentStart = index + 1;
                }
                if (segmentStart == 0) continue;
                if (segmentStart < value.Length) AppendText(segmentStart, value.Length - segmentStart);
                text.Remove();

                void AppendText(int start, int length) {
                    var segment = (TextType)text.CloneNode(false);
                    segment.Text = value.Substring(start, length);
                    text.InsertBeforeSelf(segment);
                }
            }
        }

        private static string GetLegacyDocStaticFieldInstruction(LegacyDocFieldKind fieldKind) {
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

    }
}
