namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed partial class LegacyDocDocument {
        private void AddFieldMetadataDiagnostics(LegacyDocFieldMetadataReader metadata, bool reportUnsupportedContent,
            IReadOnlyList<LegacyDocTextCharacter>? characters = null, LegacyDocFib? fib = null) {
            ISet<int>? emptyResultEnds = null;
            if (characters != null && fib.HasValue) {
                emptyResultEnds = LegacyDocField.ReportUnprojectedEmptyResults(characters, fib.Value,
                    finding => AddUnsupportedFeature(finding, reportUnsupportedContent));
            }
            foreach (LegacyDocUnsupportedFeature finding in metadata.Findings.Concat(metadata.UnprojectedLocks(emptyResultEnds))) {
                AddUnsupportedFeature(finding, reportUnsupportedContent);
            }
            if (metadata.EditedCacheCount > 0) {
                AddInfo("doc-field-cache-edit-history", $"The native DOC marks {metadata.EditedCacheCount} cached field results as edited or formatted since calculation. These edit-history flags are not byte-preserved and do not establish an invalid or stale result; no w:dirty state is inferred.");
            }
        }
    }
}
