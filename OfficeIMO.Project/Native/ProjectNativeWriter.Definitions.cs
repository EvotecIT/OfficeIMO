namespace OfficeIMO.Project;

internal sealed partial class ProjectNativeWriter {
    private void WriteDefinitions() {
        if (!ChangedTree("/Definition")) return;
        Handle("/Definition/Count");
        if (_profile == ProjectNativeProfile.Mpp8) {
            for (int index = 0; index < _document.CustomFields.Count; index++) {
                _token.ThrowIfCancellationRequested(); var definition = _document.CustomFields[index]; string path = "/Definition[" + index + "]";
                if (!uint.TryParse(definition.FieldId, out uint id)) continue;
                var field = ProjectCustomFieldIdentity.TaskFields.Concat(ProjectCustomFieldIdentity.ResourceFields).FirstOrDefault(f => f.Id == id);
                if (field == null) continue;
                Handle(path + "/FieldId");
                if (definition.FieldName == null || definition.FieldName == field.Name) Handle(path + "/FieldName");
                if (definition.Alias != null && Changed(path + "/Alias")) {
                    Handle(path + "/Alias");
                    Loss("PROJECT_NATIVE_ALIAS_LOSS", "Project 98 custom-field values are retained, but aliases have no qualified encoding and are omitted.", path + "/Alias");
                }
            }
            return;
        }
        foreach (bool task in new[] { true, false }) {
            var catalog = task ? ProjectCustomFieldIdentity.TaskFields : ProjectCustomFieldIdentity.ResourceFields;
            if (_new && !_document.CustomFields.Any(d => uint.TryParse(d.FieldId, out uint id) && catalog.Any(f => f.Id == id))) continue;
            string stream = _profile.DataRoot + "/TBknd" + (task ? "Task" : "Rsc") + "/Props";
            _file.Streams.TryGetValue(stream, out var source);
            var props = source == null ? new ProjectNativePropertySet() : new ProjectNativePropertySet(source, _token);
            var records = new Dictionary<uint, byte[]>(); byte[] trailer;
            if (source != null && ProjectNativeProperties.Read(source, _token).TryGetValue(0x04400001, out var data)) {
                int end = checked(data.Int32() + 4), count = data.Int32(8);
                if (end != 12 + count * 112 || end > data.Length || data.Int32() != data.Int32(4)) throw new InvalidDataException("Unqualified native alias envelope.");
                for (int i = 0; i < count; i++) records.Add(data.UInt32(12 + i * 112), data.Slice(12 + i * 112, 112).Copy());
                trailer = data.Slice(end, data.Length - end).Copy();
            } else {
                // Legacy field metadata has three length/count envelopes. MPP14
                // ends after two envelopes and an eight-byte empty definition header.
                trailer = new byte[_profile == ProjectNativeProfile.Mpp14 ? 32 : 36];
                Put(trailer, 0, 8); Put(trailer, 4, 8); Put(trailer, 12, 8); Put(trailer, 16, 8);
                if (_profile != ProjectNativeProfile.Mpp14) { Put(trailer, 24, 8); Put(trailer, 28, 8); }
            }
            var retained = new HashSet<string>(_document.CustomFields.Where(d => d.FieldId != null).Select(d => d.FieldId!));
            foreach (var old in Original("/Definition").Where(p => p.Key.EndsWith("/FieldId", StringComparison.Ordinal))) {
                if (!(old.Value is string text) || retained.Contains(text) || !uint.TryParse(text, out uint id) || !catalog.Any(f => f.Id == id)) continue;
                records.Remove(id);
                string path = old.Key.Substring(0, old.Key.LastIndexOf('/')); Handle(path + "/FieldId"); Handle(path + "/Alias"); Handle(path + "/FieldName");
            }
            for (int index = 0; index < _document.CustomFields.Count; index++) {
                _token.ThrowIfCancellationRequested(); var definition = _document.CustomFields[index];
                if (!uint.TryParse(definition.FieldId, out uint id)) continue;
                var field = catalog.FirstOrDefault(f => f.Id == id); if (field == null) continue;
                string path = "/Definition[" + index + "]";
                Handle(path + "/FieldId"); if (definition.FieldName == null || definition.FieldName == field.Name) Handle(path + "/FieldName");
                Handle(path + "/Alias");
                if (!Changed(path + "/Alias") && !Changed(path + "/FieldId")) continue;
                if (definition.Alias == null) { records.Remove(id); continue; }
                byte[] alias;
                try { alias = Text(definition.Alias); }
                catch (ArgumentException ex) { AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_ALIAS_VALUE", ProjectDiagnosticSeverity.Error, ex.Message, path)); continue; }
                if (alias.Length > 104) { AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_ALIAS_LENGTH", ProjectDiagnosticSeverity.Error, "Native aliases allow at most 51 UTF-16 code units.", path)); continue; }
                if (!records.TryGetValue(id, out var record)) { record = new byte[112]; Put(record, 0, unchecked((int)id)); Put(record, 4, 16); records.Add(id, record); }
                Array.Clear(record, 8, 104); Buffer.BlockCopy(alias, 0, record, 8, alias.Length);
            }
            using var buffer = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(_options.MaxOutputBytes); using var writer = new BinaryWriter(buffer);
            writer.Write(checked(8 + records.Count * 112)); writer.Write(checked(8 + records.Count * 112)); writer.Write(records.Count);
            foreach (var record in records.Values) { _token.ThrowIfCancellationRequested(); writer.Write(record); }
            writer.Write(trailer); props.Set(0x04400001, 0, buffer.ToArray());
            if (source == null && task) props.Set(0x04400000, 2, new byte[2]);
            _replacements[stream] = props.Serialize(_options.MaxOutputBytes, _token);
        }
        if (!_new) Loss("PROJECT_NATIVE_CUSTOM_REFERENCES", "Aliases are updated, but references inside opaque custom-field formulas and lookup metadata are not rewritten.", "/Definition");
    }
}
