using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryBuildModifiedMasterPersistObjects(
            PowerPointPresentation presentation, LegacyPptPackage package,
            LegacyPptProjectionMap projectionMap,
            IDictionary<uint, byte[]> rewritten) {
            SlideMasterPart[] masterParts = presentation.OpenXmlDocument
                .PresentationPart?.SlideMasterParts.ToArray()
                ?? Array.Empty<SlideMasterPart>();
            if (masterParts.Length != projectionMap.Masters.Count) return false;

            foreach (SlideMasterPart masterPart in masterParts) {
                if (!projectionMap.TryGetMaster(masterPart,
                        out LegacyPptMasterProjection? projection)
                    || projection == null) {
                    return false;
                }
                bool themeChanged = !projection.ThemeMatches(masterPart);
                if (!TryBuildMasterShapeEdits(masterPart, projection,
                        out IReadOnlyDictionary<uint, ProjectedShapeEdit>
                            shapeEdits)) {
                    return false;
                }
                if (!themeChanged && shapeEdits.Count == 0) continue;
                if (!package.PersistObjects.TryGetValue(projection.PersistId,
                        out LegacyPptPersistObject? persistObject)
                    || persistObject == null) {
                    return false;
                }
                LegacyPptRecord masterRecord = LegacyPptRecordReader.ReadSingle(
                    persistObject.RecordBytes, 0,
                    new LegacyPptImportOptions());
                byte[] masterBytes = masterRecord.CopyRecordBytes();
                if (shapeEdits.Count > 0) {
                    RecordRewrite shapeRewrite = RewriteRecord(masterRecord,
                        shapeEdits);
                    if (!shapeRewrite.Changed
                        || shapeRewrite.PatchedShapeCount != shapeEdits.Count) {
                        return false;
                    }
                    masterBytes = shapeRewrite.Bytes;
                }
                if (themeChanged) {
                    LegacyPptRecord themedRecord = LegacyPptRecordReader
                        .ReadSingle(masterBytes, 0,
                            new LegacyPptImportOptions());
                    masterBytes = LegacyPptWriter
                        .BuildPreservedMasterThemeRecord(themedRecord,
                            masterPart,
                            projection.GetChangedClassicColorSlots(masterPart));
                }
                rewritten.Add(projection.PersistId, masterBytes);
            }
            return true;
        }

        private static bool TryBuildMasterShapeEdits(SlideMasterPart masterPart,
            LegacyPptMasterProjection projection,
            out IReadOnlyDictionary<uint, ProjectedShapeEdit> edits) {
            var result = new Dictionary<uint, ProjectedShapeEdit>();
            edits = result;
            IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                .ReadMasterShapesForWrite(masterPart,
                    out string? unsupportedReason);
            if (unsupportedReason != null || shapes.Count != projection.Shapes.Count) {
                return false;
            }

            foreach (PowerPointShape shape in shapes) {
                uint? openXmlShapeId = shape.Id;
                if (!openXmlShapeId.HasValue
                    || !projection.TryGetShape(openXmlShapeId.Value,
                        out LegacyPptShapeProjection? shapeProjection)
                    || shapeProjection == null
                    || !MatchesProjectedKind(shape, shapeProjection.Kind)) {
                    return false;
                }
                LegacyPptBounds bounds = GetBounds(shape);
                LegacyPptBounds? changedBounds = BoundsEqual(
                    bounds, shapeProjection.Bounds) ? null : bounds;
                string? changedText = null;
                if (shape is PowerPointTextBox textBox) {
                    if (!MatchesProjectedTextFormatting(textBox,
                            shapeProjection)) {
                        return false;
                    }
                    string currentText = NormalizeLogicalText(textBox.Text);
                    if (!string.Equals(currentText,
                            NormalizeLogicalText(shapeProjection.Text),
                            StringComparison.Ordinal)) {
                        changedText = currentText;
                    }
                }
                if (changedBounds.HasValue || changedText != null) {
                    result.Add(shapeProjection.OfficeArtShapeId,
                        new ProjectedShapeEdit(changedBounds,
                            shapeProjection.Text, changedText,
                            interactions: null, rewriteAnimation: false,
                            animation: null));
                }
            }
            return true;
        }
    }
}
