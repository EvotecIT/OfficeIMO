using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryRewriteVbaProject(
            PowerPointPresentation presentation, LegacyPptPackage package,
            LegacyPptProjectionMap projectionMap,
            IDictionary<uint, byte[]> rewritten) {
            if (!projectionMap.VbaProject.TryGetChange(
                    presentation.OpenXmlDocument.PresentationPart,
                    out byte[]? currentProjectBytes, out bool changed)) {
                return false;
            }
            if (!changed) return true;

            uint? targetPersistId = projectionMap.VbaProject.PersistId;
            bool documentReferenceChanged = currentProjectBytes == null;
            if (currentProjectBytes != null) {
                if (!targetPersistId.HasValue) {
                    uint maximumPersistId = package.PersistObjects.Keys
                        .Concat(rewritten.Keys).DefaultIfEmpty(0U).Max();
                    if (maximumPersistId >= 0x000FFFFF) return false;
                    targetPersistId = maximumPersistId + 1U;
                    documentReferenceChanged = true;
                }
                rewritten[targetPersistId.Value] = LegacyPptWriter
                    .BuildVbaProjectStorageRecord(currentProjectBytes);
            }

            if (!documentReferenceChanged) return true;
            if (!rewritten.TryGetValue(package.DocumentPersistId,
                    out byte[]? sourceDocumentBytes)) {
                if (!package.PersistObjects.TryGetValue(
                        package.DocumentPersistId,
                        out LegacyPptPersistObject? documentObject)
                    || documentObject == null) {
                    return false;
                }
                sourceDocumentBytes = documentObject.RecordBytes;
            }
            LegacyPptRecord document = LegacyPptRecordReader.ReadSingle(
                sourceDocumentBytes!, 0, new LegacyPptImportOptions());
            rewritten[package.DocumentPersistId] = LegacyPptWriter
                .RewriteDocumentVbaInfo(document,
                    currentProjectBytes == null ? null : targetPersistId);
            return true;
        }
    }
}
