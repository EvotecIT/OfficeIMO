using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordExternalOleObjectAtom = 0x0FC3;
        private const ushort RecordExternalOleEmbed = 0x0FCC;
        private const ushort RecordExternalOleEmbedAtom = 0x0FCD;

        private static bool TryRewriteOleObjectMetadata(
            LegacyPptPackage package, byte[]? currentDocumentBytes,
            IReadOnlyList<LegacyPptOleObjectEdit> edits,
            out byte[] documentBytes) {
            documentBytes = Array.Empty<byte>();
            if (edits.Count == 0) return false;
            if (currentDocumentBytes == null) {
                if (!package.PersistObjects.TryGetValue(
                        package.DocumentPersistId,
                        out LegacyPptPersistObject? documentObject)
                    || documentObject == null) {
                    return false;
                }
                currentDocumentBytes = documentObject.RecordBytes;
            }
            LegacyPptRecord document = LegacyPptRecordReader.ReadSingle(
                currentDocumentBytes, 0, new LegacyPptImportOptions());
            var editsById = new Dictionary<uint, LegacyPptOleObjectEdit>();
            foreach (LegacyPptOleObjectEdit edit in edits) {
                if (editsById.ContainsKey(edit.Projection.Source.Id)) {
                    return false;
                }
                editsById.Add(edit.Projection.Source.Id, edit);
            }
            var matched = new HashSet<uint>();
            if (!TryRewriteOleObjectMetadataRecord(document, editsById,
                    matched, out documentBytes, out bool changed)
                || !changed || matched.Count != editsById.Count) {
                documentBytes = Array.Empty<byte>();
                return false;
            }
            return true;
        }

        private static bool TryRewriteOleObjectMetadataRecord(
            LegacyPptRecord record,
            IReadOnlyDictionary<uint, LegacyPptOleObjectEdit> editsById,
            ISet<uint> matched, out byte[] bytes, out bool changed) {
            if (record.Type == RecordExternalOleEmbed) {
                return TryRewriteOleEmbedContainer(record, editsById,
                    matched, out bytes, out changed);
            }
            if (record.Version != 0x0F || record.Children.Count == 0) {
                bytes = record.CopyRecordBytes();
                changed = false;
                return true;
            }

            var children = new List<byte[]>(record.Children.Count);
            changed = false;
            foreach (LegacyPptRecord child in record.Children) {
                if (!TryRewriteOleObjectMetadataRecord(child, editsById,
                        matched, out byte[] childBytes,
                        out bool childChanged)) {
                    bytes = record.CopyRecordBytes();
                    changed = false;
                    return false;
                }
                children.Add(childBytes);
                changed |= childChanged;
            }
            bytes = changed
                ? BuildRecord(record.Version, record.Instance, record.Type,
                    Concat(children))
                : record.CopyRecordBytes();
            return true;
        }

        private static bool TryRewriteOleEmbedContainer(
            LegacyPptRecord container,
            IReadOnlyDictionary<uint, LegacyPptOleObjectEdit> editsById,
            ISet<uint> matched, out byte[] bytes, out bool changed) {
            bytes = container.CopyRecordBytes();
            changed = false;
            if (container.Version != 0x0F || container.Instance != 0) {
                return true;
            }
            LegacyPptRecord[] objectAtoms = container.Children.Where(child =>
                child.Type == RecordExternalOleObjectAtom).ToArray();
            if (objectAtoms.Length != 1 || objectAtoms[0].PayloadLength != 24) {
                return true;
            }
            uint id = objectAtoms[0].ReadUInt32(8);
            if (!editsById.TryGetValue(id,
                    out LegacyPptOleObjectEdit? edit)) {
                return true;
            }
            if (!matched.Add(id)) return false;

            bool patchedEmbed = !edit.ColorFollowChanged;
            bool patchedObject = !edit.DrawAspectChanged
                && !edit.ProgIdChanged;
            bool patchedProgId = !edit.ProgIdChanged;
            var children = new List<byte[]>(container.Children.Count + 1);
            foreach (LegacyPptRecord child in container.Children) {
                if (child.Type == RecordExternalOleEmbedAtom
                    && edit.ColorFollowChanged) {
                    if (child.PayloadLength != 8) return false;
                    byte[] atom = child.CopyRecordBytes();
                    WriteUInt32(atom, 8, MapOleColorFollow(
                        edit.ColorFollow));
                    children.Add(atom);
                    patchedEmbed = true;
                } else if (child.Type == RecordExternalOleObjectAtom
                           && (edit.DrawAspectChanged
                               || edit.ProgIdChanged)) {
                    if (child.PayloadLength != 24) return false;
                    byte[] atom = child.CopyRecordBytes();
                    if (edit.DrawAspectChanged) {
                        WriteUInt32(atom, 8,
                            edit.ShowAsIcon ? 4U : 1U);
                    }
                    if (edit.ProgIdChanged) {
                        WriteUInt32(atom, 20, LegacyPptWriter
                            .MapOleSubType(ResolveOleProgId(edit.ProgId)));
                    }
                    children.Add(atom);
                    patchedObject = true;
                } else if (child.Type == RecordCString
                           && child.Instance == 2
                           && edit.ProgIdChanged) {
                    children.Add(LegacyPptWriter.BuildOleString(2,
                        ResolveOleProgId(edit.ProgId)));
                    patchedProgId = true;
                } else {
                    if (edit.ProgIdChanged && !patchedProgId
                        && child.Type == RecordCString
                        && child.Instance > 2) {
                        children.Add(LegacyPptWriter.BuildOleString(2,
                            ResolveOleProgId(edit.ProgId)));
                        patchedProgId = true;
                    }
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (edit.ProgIdChanged && !patchedProgId) {
                children.Add(LegacyPptWriter.BuildOleString(2,
                    ResolveOleProgId(edit.ProgId)));
                patchedProgId = true;
            }
            if (!patchedEmbed || !patchedObject || !patchedProgId) {
                return false;
            }
            bytes = BuildRecord(container.Version, container.Instance,
                container.Type, Concat(children));
            changed = true;
            return true;
        }

        private static string ResolveOleProgId(string? progId) {
            string value = string.IsNullOrWhiteSpace(progId)
                ? "Package"
                : progId!;
            if (value.IndexOf('\0') >= 0) {
                throw new InvalidDataException(
                    "The OLE ProgID contains a null character.");
            }
            return value;
        }

        private static uint MapOleColorFollow(
            DocumentFormat.OpenXml.Presentation
                .OleObjectFollowColorSchemeValues value) =>
            value == DocumentFormat.OpenXml.Presentation
                    .OleObjectFollowColorSchemeValues.Full
                ? 1U
                : value == DocumentFormat.OpenXml.Presentation
                    .OleObjectFollowColorSchemeValues.TextAndBackground
                    ? 2U
                    : 0U;
    }
}
