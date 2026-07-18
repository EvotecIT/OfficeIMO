using System.Collections.ObjectModel;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordExternalOleLink = 0x0FCE;
        private const ushort RecordExternalOleLinkAtom = 0x0FD1;

        private readonly List<LegacyPptLinkedOleObject>
            _linkedOleObjects = new();
        private readonly Dictionary<uint, LegacyPptLinkedOleObject>
            _linkedOleObjectsById = new();

        /// <summary>Gets typed linked OLE objects with exact cache storage.</summary>
        public IReadOnlyList<LegacyPptLinkedOleObject> LinkedOleObjects =>
            new ReadOnlyCollection<LegacyPptLinkedOleObject>(
                _linkedOleObjects);

        private void ParseLinkedOleObjects(LegacyPptRecord document,
            LegacyPptPackage package, LegacyPptImportOptions options) {
            foreach (LegacyPptRecord list in document.Children.Where(record =>
                         record.Type == RecordExternalObjectList)) {
                foreach (LegacyPptRecord container in list.Children.Where(
                             record => record.Type == RecordExternalOleLink)) {
                    HasLinkedOleContent = true;
                    LegacyPptLinkedOleObject? linked =
                        TryReadLinkedOleObject(container, package, options);
                    if (linked == null) continue;
                    if (_hyperlinksById.ContainsKey(linked.Id)
                        || _oleObjectsById.ContainsKey(linked.Id)
                        || _linkedOleObjectsById.ContainsKey(linked.Id)) {
                        AddDiagnostic("PPT-OLE-LINK-ID-DUPLICATE",
                            LegacyPptDiagnosticSeverity.Warning,
                            $"Linked OLE identifier {linked.Id} occurs more than once or collides with another object; later objects remain preserve-only.",
                            container.Offset);
                        continue;
                    }
                    _linkedOleObjectsById.Add(linked.Id, linked);
                    _linkedOleObjects.Add(linked);
                }
            }
        }

        private LegacyPptLinkedOleObject? TryReadLinkedOleObject(
            LegacyPptRecord container, LegacyPptPackage package,
            LegacyPptImportOptions options) {
            if (container.Version != 0x0F || container.Instance != 0) {
                AddDiagnostic("PPT-OLE-LINK-CONTAINER",
                    LegacyPptDiagnosticSeverity.Warning,
                    "An ExOleLinkContainer has an invalid record header and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord[] links = container.Children.Where(record =>
                record.Type == RecordExternalOleLinkAtom).ToArray();
            LegacyPptRecord[] objects = container.Children.Where(record =>
                record.Type == RecordExternalOleObjectAtom).ToArray();
            LegacyPptRecord[] metafiles = container.Children.Where(record =>
                record.Type == RecordMetafile).ToArray();
            if (links.Length != 1 || links[0].Version != 0
                || links[0].Instance != 0 || links[0].PayloadLength != 12
                || objects.Length != 1 || objects[0].Version != 1
                || objects[0].Instance != 0
                || objects[0].PayloadLength != 24
                || metafiles.Length > 1) {
                AddDiagnostic("PPT-OLE-LINK-ATOM",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A linked OLE object has malformed identifying atoms and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord link = links[0];
            LegacyPptRecord obj = objects[0];
            uint drawAspectValue = obj.ReadUInt32(0);
            uint type = obj.ReadUInt32(4);
            uint id = obj.ReadUInt32(8);
            uint subType = obj.ReadUInt32(12);
            uint persistId = obj.ReadUInt32(16);
            uint updateMode = link.ReadUInt32(4);
            if (type != 1 || id == 0 || persistId == 0
                || !IsOleDrawAspect(drawAspectValue)
                || updateMode is not 1 and not 3) {
                AddDiagnostic("PPT-OLE-LINK-IDENTITY",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A linked OLE object uses an invalid type, identifier, persist reference, view aspect, or update mode.",
                    obj.Offset);
                return null;
            }
            if (!TryReadOleString(container, 1, out string? menuName)
                || !TryReadOleString(container, 2, out string? progId)
                || !TryReadOleString(container, 3,
                    out string? clipboardName)) {
                AddDiagnostic("PPT-OLE-LINK-STRING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Linked OLE identifier {id} has duplicate or malformed names and remains preserve-only.",
                    container.Offset);
                return null;
            }
            if (!TryReadExternalObjectStorage(package, options, id,
                    persistId, "PPT-OLE-LINK-STORAGE", "Linked OLE",
                    obj.Offset, out byte[] storageBytes,
                    out bool compressed)) {
                return null;
            }
            return new LegacyPptLinkedOleObject(id, persistId,
                link.ReadUInt32(0), (LegacyPptOleUpdateMode)updateMode,
                (LegacyPptOleDrawAspect)drawAspectValue, subType,
                menuName, progId, clipboardName, compressed,
                metafiles.Length == 0
                    ? null
                    : metafiles[0].CopyRecordBytes(),
                storageBytes);
        }
    }
}
