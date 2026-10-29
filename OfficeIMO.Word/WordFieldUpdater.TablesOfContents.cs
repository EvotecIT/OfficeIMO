using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static void RegenerateTablesOfContents(WordDocument document, FieldEvaluationState state) {
            IReadOnlyList<WordTableOfContent> tables = WordTableOfContent.GetTablesOfContents(document);
            if (tables.Count == 0) {
                return;
            }

            var existingFields = EnumerateFields(document._wordprocessingDocument.MainDocumentPart!).ToList();
            foreach (WordTableOfContent table in tables) {
                bool locked = existingFields.Any(field => field.IsLocked
                    && WordFieldInventory.ParseInstruction(field.InstructionText).FieldType == WordFieldType.TOC
                    && field.AnchorElement.Ancestors<SdtBlock>().Contains(table.SdtBlock));
                if (locked) {
                    state.LockedTablesOfContents.Add(table.SdtBlock);
                    continue;
                }
                table.RegenerateFieldResult();
                state.RegeneratedTablesOfContents.Add(table.SdtBlock);
            }

            // The entries are final, so an open-time rebuild that OfficeIMO turned on for these tables is no longer
            // needed. A setting the caller or the loaded file enabled is kept.
            if (state.LockedTablesOfContents.Count == 0 && document.TableOfContentEnabledUpdateFieldsOnOpen) {
                document.Settings.UpdateFieldsOnOpen = false;
            }

            if (state.LockedTablesOfContents.Count == 0) document.ResetTableOfContentUpdateQueue();
        }

        private static bool IsInsideRegeneratedTableOfContents(MutableFieldCandidate candidate, FieldEvaluationState state) {
            if (state.RegeneratedTablesOfContents.Count == 0) {
                return false;
            }

            foreach (SdtBlock block in candidate.AnchorElement.Ancestors<SdtBlock>()) {
                if (state.RegeneratedTablesOfContents.Contains(block)) {
                    return true;
                }
            }

            return false;
        }
    }
}
