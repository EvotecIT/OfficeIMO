using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

internal sealed partial class ProjectNativeWriter {
    private void WriteProperties() {
        var props = new ProjectNativePropertySet(_file.Streams[_profile.Properties], _token, _profile == ProjectNativeProfile.Mpp8);
        void Property(string path, uint id, int type, Func<object, byte[]> encode) {
            Handle(path); if (!_new && !Changed(path)) return;
            _current.TryGetValue(path, out var value);
            if (value == null) props.Remove(id);
            else try { props.Set(id, type, encode(value)); }
                catch (Exception ex) when (ex is ArgumentException || ex is OverflowException) {
                    AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_PROPERTY_VALUE", ProjectDiagnosticSeverity.Error, ex.Message, path));
                }
        }
        byte[] Integer(object value) => BitConverter.GetBytes((int)value);
        byte[] Short(object value) => BitConverter.GetBytes(checked((short)(int)value));
        byte[] Date(object value) => BitConverter.GetBytes(ProjectNativeCreation.Date((DateTime)value));
        byte[] String(object value) => Text((string)value);
        Property("/Project/Name", 0x02400008, 0, String);
        byte[] ProjectGuid(Guid value) => _profile.Version <= 9
            ? System.Text.Encoding.Unicode.GetBytes(value.ToString("B").ToUpperInvariant() + "\0\0") : value.ToByteArray();
        Property("/Project/Guid", 0x02400029, 0, value => ProjectGuid((Guid)value));
        if (_new && !_document.Guid.HasValue) props.Set(0x02400029, 0, ProjectGuid(_document.NativeIdentity));
        Property("/Settings/StartDate", 0x02400002, 4, Date); Property("/Settings/FinishDate", 0x02400003, 4, Date);
        Property("/Settings/StatusDate", 0x0240003e, 4, Date);
        Property("/Settings/ScheduleFromStart", 0x02400004, 2, value => BitConverter.GetBytes((short)((bool)value ? 1 : 0)));
        Property("/Settings/MinutesPerDay", 0x0240001d, 4, Integer); Property("/Settings/MinutesPerWeek", 0x0240001e, 4, Integer);
        Property("/Settings/DaysPerMonth", 0x0240138f, 2, Short); Property("/Settings/CurrencyDigits", 0x02400012, 2, Short);
        Property("/Settings/CurrencyCode", 0x024013bb, 0, String); Property("/Settings/CurrencySymbol", 0x02400010, 0, String);
        Property("/Settings/DefaultStartTime", 0x0240001c, 2, value => BitConverter.GetBytes(checked((ushort)Exact(((TimeSpan)value).Ticks / (decimal)(TimeSpan.TicksPerMinute / 10)))));
        Property("/Settings/DefaultFinishTime", 0x02400021, 2, value => BitConverter.GetBytes(checked((ushort)Exact(((TimeSpan)value).Ticks / (decimal)(TimeSpan.TicksPerMinute / 10)))));
        Handle("/Settings/Calendar");
        if (_document.Calendar != null && _document.Calendar.IsBaseCalendar == true && !string.IsNullOrWhiteSpace(_document.Calendar.Name)) {
            props.Set(0x0240000e, 0, Text(_document.Calendar.Name ?? string.Empty));
            if (_profile.HasExtendedRecords) props.Set(0x024013c2, 0, EntityGuid(_document.Calendar, 5).ToByteArray());
        }
        else AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_CALENDAR_REQUIRED", ProjectDiagnosticSeverity.Error, "Native output requires an explicit named base calendar.", "/Settings/Calendar"));
        if (_new && !_document.Settings.StartDate.HasValue) AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_START_REQUIRED", ProjectDiagnosticSeverity.Error, "Set an explicit project start date before creating native output.", "/Settings/StartDate"));
        foreach (var table in new[] { ("Task", 1u), ("Rsc", 2u), ("Cal", 3u), ("Assn", 4u), ("Cons", 5u) }) {
            string prefix = _profile.DataRoot + "/TBknd" + table.Item1 + "/";
            if (_profile == ProjectNativeProfile.Mpp8) {
                if (_replacements.TryGetValue(prefix + "FixFix   0", out var rows)) {
                    var descriptor = ProjectNativeProperties.Read(_properties[0x02000000 | table.Item2].Copy(), _token);
                    props.Set(0x01000000 | table.Item2, 4, BitConverter.GetBytes(rows.Length / descriptor[5].Int32()));
                }
                continue;
            }
            if (!_replacements.TryGetValue(prefix + "FixedMeta", out var meta)) continue;
            props.Set(0x01000000 | table.Item2, 4, BitConverter.GetBytes(BitConverter.ToInt32(meta, 8)));
            int count = BitConverter.ToInt32(meta, 8), width = count == 0 ? 0 : (meta.Length - 16) / count, deleted = 0;
            for (int i = 0; i < count; i++) if ((meta[16 + i * width] & 2) != 0) deleted++;
            props.Set(0x00800000 | table.Item2, 4, BitConverter.GetBytes(deleted));
            props.Set(0x00010000 | table.Item2, 4, BitConverter.GetBytes(_replacements[prefix + "Var2Data"].Length));
        }
        _replacements[_profile.Properties] = props.Serialize(_options.MaxOutputBytes, _token);
        bool template = ProjectNativeProfile.IsTemplate(_options.Format);
        if (_new || template != _document.NativeInfo!.IsTemplate)
            _replacements["\u0001CompObj"] = OfficeOleCompoundObjectWriter.Write(new Guid("74b78f3a-c8c8-11d1-be11-00c04fb6faf1"),
                "Microsoft.Project 16.0", (template ? "MSProject.MPT" : "MSProject.MPP") + _profile.Version, "MSProject.Project.9");
        Metadata(OfficeOlePropertySetWriter.SummaryInformationStreamName, OfficeOlePropertySetWriter.SummaryInformationFormatId,
            new[] { ("Title", 2u), ("Subject", 3u), ("Author", 4u) });
        Metadata(OfficeOlePropertySetWriter.DocumentSummaryInformationStreamName, OfficeOlePropertySetWriter.DocumentSummaryInformationFormatId,
            new[] { ("Manager", 14u), ("Company", 15u) });
    }
    private void Metadata(string stream, Guid section, (string Name, uint Id)[] fields) {
        var changes = new Dictionary<uint, string?>();
        if (_new && section == OfficeOlePropertySetWriter.SummaryInformationFormatId)
            changes[18] = ProjectNativeSource.ProducerMarker + ":" + _document.NativeIdentity.ToString("D");
        foreach (var field in fields) {
            string path = "/Project/" + field.Name; Handle(path);
            if (!_new && !Changed(path)) continue;
            _current.TryGetValue(path, out var value);
            if (value is string text) {
                try { _ = Text(text); }
                catch (ArgumentException ex) { AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_METADATA_VALUE", ProjectDiagnosticSeverity.Error, ex.Message, path)); continue; }
            }
            changes.Add(field.Id, (string?)value);
        }
        if (changes.Count == 0) return;
        _file.Streams.TryGetValue(stream, out var source);
        _replacements[stream] = OfficeOlePropertySetEditor.RewriteStrings(source, section, changes, _token);
    }
}
