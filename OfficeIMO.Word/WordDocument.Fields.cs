namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Inspects simple and complex Word fields across the main document, headers, footers, footnotes and endnotes.
        /// </summary>
        /// <returns>A deterministic read-only field inventory with raw instructions, parsed tokens, result text and location metadata.</returns>
        public IReadOnlyList<WordFieldInfo> InspectFields() {
            return WordFieldInventory.Inspect(this);
        }

        /// <summary>
        /// Updates deterministic field results supported by OfficeIMO and returns structured per-field diagnostics.
        /// </summary>
        /// <returns>A report describing updated, skipped, unsupported and malformed fields.</returns>
        public WordFieldUpdateReport UpdateFieldsAndGetReport() {
            return WordFieldUpdater.Update(this, WordFieldUpdateOptions.Default);
        }

        /// <summary>
        /// Updates deterministic field results supported by OfficeIMO and returns structured per-field diagnostics.
        /// </summary>
        /// <param name="options">Options controlling deterministic field refresh behavior.</param>
        /// <returns>A report describing updated, skipped, unsupported and malformed fields.</returns>
        public WordFieldUpdateReport UpdateFieldsAndGetReport(WordFieldUpdateOptions options) {
            return WordFieldUpdater.Update(this, options ?? WordFieldUpdateOptions.Default);
        }
    }
}
