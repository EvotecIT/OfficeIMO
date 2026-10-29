using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static void RegenerateTablesOfContents(WordDocument document, FieldEvaluationState state) {
            IReadOnlyList<WordTableOfContent> tables = WordTableOfContent.GetTablesOfContents(document);
            if (tables.Count == 0) {
                return;
            }

            foreach (WordTableOfContent table in tables) {
                table.RegenerateFieldResult();
                state.RegeneratedTablesOfContents.Add(table.SdtBlock);
            }

            // The entries are final, so an open-time rebuild that OfficeIMO turned on for these tables is no longer
            // needed. A setting the caller or the loaded file enabled is kept.
            if (document.TableOfContentEnabledUpdateFieldsOnOpen) {
                document.Settings.UpdateFieldsOnOpen = false;
            }

            document.ResetTableOfContentUpdateQueue();
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
