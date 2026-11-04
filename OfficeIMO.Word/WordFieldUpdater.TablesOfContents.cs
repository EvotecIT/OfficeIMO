using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static void RegenerateTablesOfContents(WordDocument document, FieldEvaluationState state) {
            IReadOnlyList<WordTableOfContent> tables = WordTableOfContent.GetTablesOfContents(document);
            if (tables.Count == 0) {
                return;
            }

            bool anyLocked = false;
            foreach (WordTableOfContent table in tables) {
                bool locked = state.LockedTocFields.Any(field => field.AnchorElement.Ancestors<SdtBlock>().Contains(table.SdtBlock));
                if (locked) {
                    anyLocked = true;
                    continue;
                }
                table.RegenerateFieldResult();
                state.RegeneratedTablesOfContents.Add(table.SdtBlock);
            }

            // The entries are final, so an open-time rebuild that OfficeIMO turned on for these tables is no longer
            // needed. A setting the caller or the loaded file enabled is kept.
            if (!anyLocked && document.TableOfContentEnabledUpdateFieldsOnOpen) {
                document.Settings.UpdateFieldsOnOpen = false;
            }

            if (!anyLocked) document.ResetTableOfContentUpdateQueue();
        }

        private static bool IsLockedToc(MutableFieldCandidate field) => field.IsLocked
            && WordFieldInventory.ParseInstruction(field.InstructionText).FieldType == WordFieldType.TOC;

        private static bool IsInsideLockedToc(MutableFieldCandidate candidate, IEnumerable<MutableFieldCandidate> lockedFields) {
            foreach (MutableFieldCandidate locked in lockedFields) {
                if (!string.Equals(candidate.PartUri, locked.PartUri, StringComparison.Ordinal)) continue;
                if (ReferenceEquals(candidate.AnchorElement, locked.AnchorElement)) return true;
                if (locked.SimpleField != null && candidate.AnchorElement.Ancestors<SimpleField>().Contains(locked.SimpleField)) return true;
                if (locked.EndRun != null && candidate.NestingLevel > locked.NestingLevel
                    && candidate.AnchorElement.IsAfter(locked.AnchorElement)
                    && (ReferenceEquals(candidate.AnchorElement, locked.EndRun) || candidate.AnchorElement.IsBefore(locked.EndRun))) return true;
            }
            return false;
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
