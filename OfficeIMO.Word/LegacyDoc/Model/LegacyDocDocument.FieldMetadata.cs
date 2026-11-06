namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed partial class LegacyDocDocument {
        private void AddFieldMetadataDiagnostics(LegacyDocFieldMetadataReader metadata, bool reportUnsupportedContent) {
            foreach (LegacyDocUnsupportedFeature finding in metadata.Findings.Concat(metadata.UnprojectedLocks())) {
                AddUnsupportedFeature(finding, reportUnsupportedContent);
            }
            if (metadata.EditedCacheCount > 0) {
                AddInfo("doc-field-cache-edit-history", $"The native DOC marks {metadata.EditedCacheCount} cached field results as edited or formatted since calculation. These edit-history flags are not byte-preserved and do not establish an invalid or stale result; no w:dirty state is inferred.");
            }
        }
    }
}
