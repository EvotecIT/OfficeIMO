using System.Collections.ObjectModel;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordExternalOleObjectAtom = 0x0FC3;
        private const ushort RecordExternalOleEmbed = 0x0FCC;
        private const ushort RecordExternalOleEmbedAtom = 0x0FCD;
        private const ushort RecordExternalObjectRefAtom = 0x0BC1;
        private const ushort RecordMetafile = 0x0FC1;

        private readonly List<LegacyPptEmbeddedOleObject> _oleObjects = new();
        private readonly Dictionary<uint, LegacyPptEmbeddedOleObject>
            _oleObjectsById = new();
        private readonly HashSet<uint> _referencedExternalObjectIds = new();

        /// <summary>Gets decoded embedded OLE objects by document identifier.</summary>
        public IReadOnlyList<LegacyPptEmbeddedOleObject> OleObjects =>
            new ReadOnlyCollection<LegacyPptEmbeddedOleObject>(_oleObjects);

        private void ParseOleObjects(LegacyPptRecord document,
            LegacyPptPackage package, LegacyPptImportOptions options) {
            foreach (LegacyPptRecord list in document.Children.Where(record =>
                         record.Type == RecordExternalObjectList)) {
                foreach (LegacyPptRecord container in list.Children.Where(record =>
                             record.Type == RecordExternalOleEmbed)) {
                    HasEmbeddedOleContent = true;
                    LegacyPptEmbeddedOleObject? ole = TryReadEmbeddedOleObject(
                        container, package, options);
                    if (ole == null) continue;
                    if (_hyperlinksById.ContainsKey(ole.Id)
                        || _oleObjectsById.ContainsKey(ole.Id)) {
                        AddDiagnostic("PPT-OLE-ID-DUPLICATE",
                            LegacyPptDiagnosticSeverity.Warning,
                            $"Embedded OLE identifier {ole.Id} occurs more than once; later objects remain preserve-only.",
                            container.Offset);
                        continue;
                    }
                    _oleObjectsById.Add(ole.Id, ole);
                    _oleObjects.Add(ole);
                }
            }
            ParseLinkedOleObjects(document, package, options);
            ParseActiveXControls(document, package, options);
            ParseMediaObjects(document, options);
        }

        private LegacyPptEmbeddedOleObject? TryReadEmbeddedOleObject(
            LegacyPptRecord container, LegacyPptPackage package,
            LegacyPptImportOptions options) {
            if (container.Version != 0x0F || container.Instance != 0) {
                AddDiagnostic("PPT-OLE-CONTAINER",
                    LegacyPptDiagnosticSeverity.Warning,
                    "An ExOleEmbedContainer has an invalid record header and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord[] embeds = container.Children.Where(record =>
                record.Type == RecordExternalOleEmbedAtom).ToArray();
            LegacyPptRecord[] objects = container.Children.Where(record =>
                record.Type == RecordExternalOleObjectAtom).ToArray();
            if (embeds.Length != 1 || embeds[0].Version != 0
                || embeds[0].Instance != 0 || embeds[0].PayloadLength != 8
                || objects.Length != 1 || objects[0].Version != 1
                || objects[0].Instance != 0 || objects[0].PayloadLength != 24) {
                AddDiagnostic("PPT-OLE-ATOM",
                    LegacyPptDiagnosticSeverity.Warning,
                    "An embedded OLE object has malformed identifying atoms and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord embed = embeds[0];
            LegacyPptRecord obj = objects[0];
            uint drawAspectValue = obj.ReadUInt32(0);
            uint type = obj.ReadUInt32(4);
            uint id = obj.ReadUInt32(8);
            uint subType = obj.ReadUInt32(12);
            uint persistId = obj.ReadUInt32(16);
            uint colorFollowValue = embed.ReadUInt32(0);
            if (type != 0 || id == 0 || persistId == 0
                || !IsOleDrawAspect(drawAspectValue)
                || colorFollowValue > 2) {
                AddDiagnostic("PPT-OLE-IDENTITY",
                    LegacyPptDiagnosticSeverity.Warning,
                    "An embedded OLE object uses an invalid type, identifier, persist reference, view aspect, or color-follow value.",
                    obj.Offset);
                return null;
            }
            if (!TryReadOleString(container, 1, out string? menuName)
                || !TryReadOleString(container, 2, out string? progId)
                || !TryReadOleString(container, 3, out string? clipboardName)) {
                AddDiagnostic("PPT-OLE-STRING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Embedded OLE identifier {id} has duplicate or malformed names and remains preserve-only.",
                    container.Offset);
                return null;
            }
            if (!TryReadExternalObjectStorage(package, options, id,
                    persistId, "PPT-OLE-STORAGE", "Embedded OLE",
                    obj.Offset, out byte[] storageBytes,
                    out bool compressed)) {
                return null;
            }
            return new LegacyPptEmbeddedOleObject(id, persistId,
                (LegacyPptOleDrawAspect)drawAspectValue, subType,
                (LegacyPptOleColorFollow)colorFollowValue,
                embed.ReadByte(4) != 0, embed.ReadByte(5) != 0,
                embed.ReadByte(6) != 0, menuName, progId, clipboardName,
                compressed, storageBytes);
        }

        private static bool TryReadOleString(LegacyPptRecord container,
            ushort instance, out string? value) {
            value = null;
            LegacyPptRecord[] strings = container.Children.Where(record =>
                record.Type == RecordCString && record.Instance == instance)
                .ToArray();
            if (strings.Length == 0) return true;
            if (strings.Length != 1 || strings[0].Version != 0
                || (strings[0].PayloadLength & 1) != 0) return false;
            value = strings[0].ReadUtf16Text().TrimEnd('\0');
            return true;
        }

        private void ReadShapeExternalObject(
            LegacyPptRecord shapeContainer, LegacyPptImportOptions options,
            out LegacyPptEmbeddedOleObject? embedded,
            out LegacyPptLinkedOleObject? linked,
            out LegacyPptActiveXControl? activeX,
            out LegacyPptMedia? media) {
            embedded = null;
            linked = null;
            activeX = null;
            media = null;
            LegacyPptRecord[] references = shapeContainer.Children
                .Where(record => record.Type == OfficeArtClientData)
                .SelectMany(record => record.Children)
                .Where(record => record.Type == RecordExternalObjectRefAtom)
                .ToArray();
            if (references.Length == 0) return;
            if (references.Length != 1 || references[0].Version != 0
                || references[0].Instance != 0
                || references[0].PayloadLength != 4) {
                AddDiagnostic("PPT-OLE-SHAPE-REFERENCE",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A shape has duplicate or malformed ExObjRefAtom records and remains preserve-only.",
                    shapeContainer.Offset);
                return;
            }
            uint id = references[0].ReadUInt32(0);
            bool found = _oleObjectsById.TryGetValue(id, out embedded)
                || _linkedOleObjectsById.TryGetValue(id, out linked)
                || _activeXControlsById.TryGetValue(id, out activeX)
                || _mediaById.TryGetValue(id, out media);
            if (!found) {
                if (options.ReportUnsupportedContent) {
                    AddDiagnostic("PPT-OLE-SHAPE-TARGET",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"A shape references unavailable external object {id} and remains preserve-only.",
                        references[0].Offset);
                }
                return;
            }
            if (!_referencedExternalObjectIds.Add(id)) {
                AddDiagnostic("PPT-OLE-SHAPE-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"External object identifier {id} is referenced by more than one shape; later references remain preserve-only.",
                    references[0].Offset);
                embedded = null;
                linked = null;
                activeX = null;
                media = null;
                return;
            }
        }

        private bool TryReadExternalObjectStorage(LegacyPptPackage package,
            LegacyPptImportOptions options, uint id, uint persistId,
            string diagnosticCode, string objectName, long? offset,
            out byte[] storageBytes, out bool compressed) {
            storageBytes = Array.Empty<byte>();
            compressed = false;
            string? storageReason = null;
            string? compoundReason = null;
            bool valid = package.PersistObjects.TryGetValue(persistId,
                    out LegacyPptPersistObject? persistObject)
                && LegacyPptOleStorageCodec.TryDecode(persistObject!, options,
                    _recordBudget, _decodedStorageBudget, out storageBytes,
                    out compressed,
                    out storageReason)
                && LegacyPptCompoundStorageValidator.TryRead(storageBytes,
                    options,
                    out OfficeCompoundFile? compound, out compoundReason)
                && compound != null;
            if (valid) return true;
            AddDiagnostic(diagnosticCode,
                LegacyPptDiagnosticSeverity.Warning,
                $"{objectName} identifier {id} has no valid compound storage and remains preserve-only: "
                + (storageReason ?? compoundReason ?? "missing persist object"),
                offset);
            storageBytes = Array.Empty<byte>();
            compressed = false;
            return false;
        }

        private void ValidateExternalObjectSlideReferences() {
            var slideIds = new HashSet<uint>(_slides.Select(slide =>
                slide.SlideId));
            foreach (LegacyPptLinkedOleObject linked in _linkedOleObjects) {
                if (linked.SlideId != 0
                    && !slideIds.Contains(linked.SlideId)) {
                    AddDiagnostic("PPT-OLE-LINK-SLIDE",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"Linked OLE identifier {linked.Id} references missing slide {linked.SlideId}; the object remains preserve-only.",
                        null);
                }
            }
            foreach (LegacyPptActiveXControl control in _activeXControls) {
                if (control.SlideId != 0
                    && !slideIds.Contains(control.SlideId)) {
                    AddDiagnostic("PPT-ACTIVEX-SLIDE",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"ActiveX identifier {control.Id} references missing slide {control.SlideId}; the control remains preserve-only.",
                        null);
                }
            }
        }

        private void ValidateExternalObjectIdSeed(LegacyPptRecord document,
            LegacyPptImportOptions options) {
            if (!options.ReportUnsupportedContent) return;
            uint greatestId = _hyperlinks.Select(item => item.Id)
                .Concat(_oleObjects.Select(item => item.Id))
                .Concat(_linkedOleObjects.Select(item => item.Id))
                .Concat(_activeXControls.Select(item => item.Id))
                .Concat(_media.Select(item => item.Id))
                .DefaultIfEmpty(0U).Max();
            foreach (LegacyPptRecord list in document.Children.Where(record =>
                         record.Type == RecordExternalObjectList)) {
                LegacyPptRecord[] atoms = list.Children.Where(record =>
                    record.Type == RecordExternalObjectListAtom).ToArray();
                if (atoms.Length == 1 && atoms[0].Version == 0
                    && atoms[0].Instance == 0
                    && atoms[0].PayloadLength == 4
                    && atoms[0].ReadUInt32(0) < greatestId) {
                    AddDiagnostic("PPT-EXTERNAL-OBJECT-ID-SEED",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"The external-object id seed {atoms[0].ReadUInt32(0)} is below live identifier {greatestId}; new objects require a repaired seed.",
                        atoms[0].Offset);
                }
            }
        }

        private static bool IsOleDrawAspect(uint value) => value == 1
            || value == 2 || value == 4 || value == 8;
    }
}
