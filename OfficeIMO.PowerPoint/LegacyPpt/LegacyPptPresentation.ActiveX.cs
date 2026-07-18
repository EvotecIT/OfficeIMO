using System.Collections.ObjectModel;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordExternalOleControl = 0x0FEE;
        private const ushort RecordExternalOleControlAtom = 0x0FFB;
        private readonly List<LegacyPptActiveXControl>
            _activeXControls = new();
        private readonly Dictionary<uint, LegacyPptActiveXControl>
            _activeXControlsById = new();

        /// <summary>Gets typed ActiveX controls with exact Office Forms storage.</summary>
        public IReadOnlyList<LegacyPptActiveXControl> ActiveXControls =>
            new ReadOnlyCollection<LegacyPptActiveXControl>(
                _activeXControls);

        private void ParseActiveXControls(LegacyPptRecord document,
            LegacyPptPackage package, LegacyPptImportOptions options) {
            foreach (LegacyPptRecord list in document.Children.Where(record =>
                         record.Type == RecordExternalObjectList)) {
                foreach (LegacyPptRecord container in list.Children.Where(
                             record => record.Type ==
                                 RecordExternalOleControl)) {
                    HasActiveXContent = true;
                    LegacyPptActiveXControl? control = TryReadActiveXControl(
                        container, package, options);
                    if (control == null) continue;
                    if (_hyperlinksById.ContainsKey(control.Id)
                        || _oleObjectsById.ContainsKey(control.Id)
                        || _linkedOleObjectsById.ContainsKey(control.Id)
                        || _activeXControlsById.ContainsKey(control.Id)) {
                        AddDiagnostic("PPT-ACTIVEX-ID-DUPLICATE",
                            LegacyPptDiagnosticSeverity.Warning,
                            $"ActiveX identifier {control.Id} occurs more than once or collides with another object; later controls remain preserve-only.",
                            container.Offset);
                        continue;
                    }
                    _activeXControlsById.Add(control.Id, control);
                    _activeXControls.Add(control);
                }
            }
        }

        private LegacyPptActiveXControl? TryReadActiveXControl(
            LegacyPptRecord container, LegacyPptPackage package,
            LegacyPptImportOptions options) {
            if (container.Version != 0x0F || container.Instance != 0) {
                AddDiagnostic("PPT-ACTIVEX-CONTAINER",
                    LegacyPptDiagnosticSeverity.Warning,
                    "An ExControlContainer has an invalid record header and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord[] controls = container.Children.Where(record =>
                record.Type == RecordExternalOleControlAtom).ToArray();
            LegacyPptRecord[] objects = container.Children.Where(record =>
                record.Type == RecordExternalOleObjectAtom).ToArray();
            LegacyPptRecord[] metafiles = container.Children.Where(record =>
                record.Type == RecordMetafile).ToArray();
            if (controls.Length != 1 || controls[0].Version != 0
                || controls[0].Instance != 0
                || controls[0].PayloadLength != 4
                || objects.Length != 1 || objects[0].Version != 1
                || objects[0].Instance != 0
                || objects[0].PayloadLength != 24
                || metafiles.Length > 1) {
                AddDiagnostic("PPT-ACTIVEX-ATOM",
                    LegacyPptDiagnosticSeverity.Warning,
                    "An ActiveX control has malformed identifying records and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord control = controls[0];
            LegacyPptRecord obj = objects[0];
            uint drawAspectValue = obj.ReadUInt32(0);
            uint type = obj.ReadUInt32(4);
            uint id = obj.ReadUInt32(8);
            uint subType = obj.ReadUInt32(12);
            uint persistId = obj.ReadUInt32(16);
            if (type != 2 || id == 0 || persistId == 0
                || !IsOleDrawAspect(drawAspectValue)) {
                AddDiagnostic("PPT-ACTIVEX-IDENTITY",
                    LegacyPptDiagnosticSeverity.Warning,
                    "An ActiveX control uses an invalid type, identifier, persist reference, or view aspect.",
                    obj.Offset);
                return null;
            }
            if (!TryReadOleString(container, 1, out string? menuName)
                || !TryReadOleString(container, 2, out string? progId)
                || !TryReadOleString(container, 3,
                    out string? clipboardName)) {
                AddDiagnostic("PPT-ACTIVEX-STRING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"ActiveX identifier {id} has duplicate or malformed names and remains preserve-only.",
                    container.Offset);
                return null;
            }
            if (!TryReadExternalObjectStorage(package, options, id,
                    persistId, "PPT-ACTIVEX-STORAGE", "ActiveX",
                    obj.Offset, out byte[] storageBytes,
                    out bool compressed)) {
                return null;
            }
            return new LegacyPptActiveXControl(id, persistId,
                control.ReadUInt32(0),
                (LegacyPptOleDrawAspect)drawAspectValue, subType,
                menuName, progId, clipboardName, compressed,
                metafiles.Length == 0
                    ? null
                    : metafiles[0].CopyRecordBytes(),
                storageBytes);
        }
    }
}
