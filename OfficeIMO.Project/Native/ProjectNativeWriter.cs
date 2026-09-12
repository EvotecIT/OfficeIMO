using OfficeIMO.Core.Internal;
using System.Text;

namespace OfficeIMO.Project;

/// <summary>Plans native updates against immutable source bytes and a complete typed model inventory.</summary>
internal sealed partial class ProjectNativeWriter {
    internal const byte SyntheticProjectSummaryKind = byte.MaxValue;
    private readonly ProjectDocument _document;
    private readonly ProjectSaveOptions _options;
    private readonly CancellationToken _token;
    private readonly OfficeCompoundFile _file;
    private readonly Dictionary<uint, ProjectNativeValue> _properties;
    private readonly Dictionary<string, byte[]> _replacements = new Dictionary<string, byte[]>();
    private readonly Dictionary<string, object?> _current, _original;
    private readonly HashSet<string> _handled = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _changes = new HashSet<string>(StringComparer.Ordinal), _changedTrees = new HashSet<string>(StringComparer.Ordinal);
    private readonly Dictionary<string, List<string>> _keys = new Dictionary<string, List<string>>(StringComparer.Ordinal);
    private readonly List<ProjectDiagnostic> _diagnostics = new List<ProjectDiagnostic>();
    private readonly bool _new;
    private readonly ProjectNativeProfile _profile;
    private readonly bool _structureChanged, _scheduleChanged, _projectGuidChanged, _taskGuidsChanged, _resourceGuidsChanged, _calendarGuidsChanged, _calendarBindingsChanged;

    private ProjectNativeWriter(ProjectDocument document, ProjectSaveOptions options, CancellationToken token) {
        _document = document; _options = options; _token = token; _profile = ProjectNativeProfile.ForFormat(options.Format);
        _new = document.NativeSource == null || document.NativeInfo!.Profile != _profile;
        _current = ProjectModelSnapshot.Capture(document, token);
        _original = _new ? new Dictionary<string, object?>() : document.NativeSource!.Snapshot;
        foreach (string key in _current.Keys.Concat(_original.Keys.Where(k => !_current.ContainsKey(k)))) {
            token.ThrowIfCancellationRequested(); string group = Group(key);
            if (!_keys.TryGetValue(group, out var keys)) _keys.Add(group, keys = new List<string>());
            keys.Add(key); _current.TryGetValue(key, out var current); _original.TryGetValue(key, out var original);
            if (Equals(current, original)) continue;
            _changes.Add(key);
            for (int i = 1; i < key.Length; i++) if (key[i] == '/' || key[i] == '[') _changedTrees.Add(key.Substring(0, i));
        }
        _structureChanged = _changes.Any(k => k.EndsWith("/Uid", StringComparison.Ordinal) || k.EndsWith("/Guid", StringComparison.Ordinal) || k.EndsWith("/Parent", StringComparison.Ordinal) || k.EndsWith("/Position", StringComparison.Ordinal)
            || k.StartsWith("/Dependency", StringComparison.Ordinal) || k.EndsWith("/Task", StringComparison.Ordinal) || k.EndsWith("/Resource", StringComparison.Ordinal));
        _scheduleChanged = _changes.Any(k => !k.StartsWith("/Project/", StringComparison.Ordinal)
            && !k.EndsWith("/Name", StringComparison.Ordinal) && !k.EndsWith("/DisplayId", StringComparison.Ordinal) && !k.EndsWith("/Guid", StringComparison.Ordinal)
            && !k.EndsWith("/Initials", StringComparison.Ordinal) && !k.EndsWith("/Group", StringComparison.Ordinal) && !k.EndsWith("/EmailAddress", StringComparison.Ordinal)
            && !k.EndsWith("/Notes", StringComparison.Ordinal) && !k.EndsWith("/Wbs", StringComparison.Ordinal) && !k.EndsWith("/Contact", StringComparison.Ordinal));
        _projectGuidChanged = _changes.Contains("/Project/Guid");
        _taskGuidsChanged = _changes.Any(k => k.StartsWith("/Task[", StringComparison.Ordinal) && k.EndsWith("/Guid", StringComparison.Ordinal));
        _resourceGuidsChanged = _changes.Any(k => k.StartsWith("/Resource[", StringComparison.Ordinal) && k.EndsWith("/Guid", StringComparison.Ordinal));
        _calendarGuidsChanged = _changes.Any(k => k.StartsWith("/Calendar[", StringComparison.Ordinal) && k.EndsWith("/Guid", StringComparison.Ordinal));
        _calendarBindingsChanged = _changes.Any(k => k.EndsWith("/Calendar", StringComparison.Ordinal)
            || k.StartsWith("/Resource[", StringComparison.Ordinal) && k.EndsWith("/Uid", StringComparison.Ordinal));
        if (_new) {
            var stagingStart = new DateTime(2000, 1, 3, 8, 0, 0);
            if (document.Settings.StartDate.HasValue) {
                try { _ = ProjectNativeCreation.Date(document.Settings.StartDate.Value); stagingStart = document.Settings.StartDate.Value; }
                catch (ArgumentException) { /* WriteProperties reports the unrepresentable model value before output. */ }
            }
            _file = ProjectNativeCreation.Empty(stagingStart, options.MaxOutputBytes, token, _profile);
        }
        else {
            var limits = new OfficeCompoundReadOptions(int.MaxValue, int.MaxValue, options.MaxOutputBytes, options.MaxOutputBytes);
            if (!OfficeCompoundFileReader.TryRead(document.NativeSource!.Bytes, limits, token, out var file, out var error) || file == null)
                throw new InvalidDataException("Native source could not be staged: " + error);
            _file = file;
        }
        _properties = ProjectNativeProperties.Read(_file.Streams[_profile.Properties], token, _profile == ProjectNativeProfile.Mpp8);
    }

    internal static (ProjectReport Report, byte[]? Bytes) Plan(ProjectDocument document, ProjectSaveOptions options, bool serialize, CancellationToken token) {
        var writer = new ProjectNativeWriter(document, options, token);
        try { writer.Build(); }
        catch (Exception ex) when (ex is ArgumentException || ex is OverflowException || ex is NotSupportedException) {
            writer.AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_VALUE_UNREPRESENTABLE", ProjectDiagnosticSeverity.Error, ex.Message, "/"));
        }
        var report = new ProjectReport(document.Revision, writer._diagnostics);
        if (!serialize || report.HasErrors || (report.HasLoss && options.LossPolicy == OfficeConversionLossPolicy.Block)) return (report, null);
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(writer._file, writer._replacements, maxOutputBytes: options.MaxOutputBytes, cancellationToken: token);
        return (report, bytes);
    }

    private void Build() {
        ValidateIdentities();
        WriteTasks(); WriteResources(); WriteAssignments(); WriteCalendars(); WriteDependencies(); WriteDefinitions(); WriteProperties();
        foreach (var prior in _document.NativeSource?.UnrepresentedValues ?? Array.Empty<ProjectDiagnostic>()) {
            bool present = _current.ContainsKey(prior.Location) || _keys.TryGetValue(Group(prior.Location), out var keys) &&
                keys.Any(k => _current.ContainsKey(k) && (k.StartsWith(prior.Location + "/", StringComparison.Ordinal) || k.StartsWith(prior.Location + "[", StringComparison.Ordinal)));
            if (!present) Handle(prior.Location);
            else if (!Changed(prior.Location) && !ChangedTree(prior.Location)) AddDiagnostic(prior);
        }
        foreach (string key in _changes) {
            _token.ThrowIfCancellationRequested();
            if (_handled.Contains(key)) continue;
            AddDiagnostic(new ProjectDiagnostic(_new ? "PROJECT_NATIVE_FIELD_LOSS" : "PROJECT_NATIVE_EDIT_UNSUPPORTED",
                _new ? ProjectDiagnosticSeverity.Warning : ProjectDiagnosticSeverity.Error,
                "This modeled value has no qualified native writer. It cannot be represented by this operation.", key, true));
        }
        if (!_new && StructureChanged) Loss("PROJECT_NATIVE_STRUCTURAL_OPAQUE", "Structural edits retain opaque presentation, calculation, and auxiliary records. References inside those records are not remapped.", "/");
        if (!_new && ScheduleChanged) Loss("PROJECT_NATIVE_STORED_TOTALS", "Mapped schedule edits retain producer work/cost curves and calculated totals. Recalculate and verify them in Project.", "/");
        if (!_new && _document.NativeInfo!.HasSignatureStorage) Loss("PROJECT_NATIVE_SIGNATURE_INVALIDATED", "Changing signed native content invalidates its existing signature; the signature is not renewed.", "/");
        if (_new && _document.NativeSource != null)
            Loss("PROJECT_NATIVE_GENERATION_LOSS", "Conversion creates the selected native generation from modeled values. Unmodeled source records, presentation, macros, signatures, and embedded content are omitted.", "/");
        if (_new && _document.Source != null && _document.ReadDiagnostics.Any(d => d.RepresentsLoss || d.Code.Contains("PRESERVED")))
            Loss("PROJECT_XML_EXTENSION_LOSS", "XML extension content has no native representation and is omitted.", "/");
    }

    private void ValidateIdentities() {
        foreach (IEnumerable<ProjectEntity> group in new IEnumerable<ProjectEntity>[] {
            _document.TaskIndex.Values, _document.ResourceIndex.Values, _document.CalendarIndex.Values, _document.AssignmentIndex.Values }) {
            var identities = new HashSet<Guid>();
            foreach (var entity in group) {
                _token.ThrowIfCancellationRequested();
                if (entity.Guid.HasValue && (entity.Guid.Value == Guid.Empty || !identities.Add(entity.Guid.Value)))
                    AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_GUID_INVALID", ProjectDiagnosticSeverity.Error,
                        "Native entity GUIDs must be nonempty and unique within their entity kind.", "/"));
            }
        }
    }
    private void Loss(string code, string message, string path) => AddDiagnostic(new ProjectDiagnostic(code, ProjectDiagnosticSeverity.Warning, message, path, true));
    private void AddDiagnostic(ProjectDiagnostic diagnostic) {
        if (_diagnostics.Count < 1000) _diagnostics.Add(diagnostic);
        else if (_diagnostics.Count == 1000) _diagnostics.Add(new ProjectDiagnostic("PROJECT_NATIVE_DIAGNOSTICS_TRUNCATED",
            ProjectDiagnosticSeverity.Error, "Native assessment exceeded its diagnostic budget; additional findings are omitted.", "/"));
    }
    private bool Changed(string key) => _changes.Contains(key);
    private bool StructureChanged => _structureChanged;
    private bool ScheduleChanged => _scheduleChanged;
    private void Handle(string key) { if (_changes.Contains(key)) _handled.Add(key); }
    private void HandleTree(string prefix) { foreach (string key in Keys(prefix)) Handle(key); }
    private bool ChangedTree(string prefix) => _changedTrees.Contains(prefix);
    private static string Group(string path) { int next = path.IndexOf('/', 1); return next < 0 ? path : path.Substring(0, next); }
    private IEnumerable<string> Keys(string prefix) {
        string group = Group(prefix);
        IEnumerable<string> candidates = _keys.TryGetValue(group, out var keys) ? keys : Enumerable.Empty<string>();
        if (group.IndexOf('[') < 0) candidates = candidates.Concat(_keys.Where(p => p.Key.StartsWith(group + "[", StringComparison.Ordinal)).SelectMany(p => p.Value));
        return candidates.Where(k => k.StartsWith(prefix + "/", StringComparison.Ordinal) || k.StartsWith(prefix + "[", StringComparison.Ordinal));
    }
    private IEnumerable<KeyValuePair<string, object?>> Original(string prefix) {
        foreach (string key in Keys(prefix)) if (_original.TryGetValue(key, out var value)) yield return new KeyValuePair<string, object?>(key, value);
    }
    private static string Path(ProjectEntity entity, string kind) => "/" + kind + "[UID=" + entity.Uid + "]";
    private IProjectNativeTableEditor Editor(string name, uint id, uint uid) => _profile == ProjectNativeProfile.Mpp8
        ? new ProjectNativeLegacy8Editor(_file, _properties, name, id, uid, _options.MaxOutputBytes, _token)
        : new ProjectNativeTableEditor(_file, _properties, name, id, uid, _options.MaxOutputBytes, _token, _profile);
    private void Identity(IProjectNativeTableEditor editor, int uid, uint id, Guid value) {
        if (_profile.HasExtendedRecords) editor.Set(uid, id, value.ToByteArray());
    }
    private void SortPosition(IProjectNativeTableEditor editor, int uid, uint id, int position) {
        if (_profile.HasExtendedRecords) editor.Set(uid, id, BitConverter.GetBytes((double)position));
    }
    private void Removed(IProjectNativeTableEditor editor, string kind, IEnumerable<int> retained, bool preserveZero = false) {
        var keep = new HashSet<int>(retained);
        foreach (int uid in editor.Uids.ToArray()) if (!keep.Contains(uid) && (!preserveZero || uid != 0)) { editor.Delete(uid); HandleTree("/" + kind + "[UID=" + uid + "]"); }
    }
    private void Field(IProjectNativeTableEditor editor, int uid, string path, string name, uint id, Kind kind = Kind.Integer, decimal scale = 1, uint format = 0, bool force = false) {
        string key = path + "/" + name;
        if (!editor.HasField(id)) {
            if (!_current.TryGetValue(key, out var absent) || absent == null) Handle(key);
            return; // The inventory reports non-null values absent from this generation's map.
        }
        Handle(key);
        if (!force && !Changed(key)) return;
        _current.TryGetValue(key, out var value);
        if (value == null) {
            if (kind == Kind.Guid) AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_GUID_REQUIRED", ProjectDiagnosticSeverity.Error, "An existing native relationship identity cannot be cleared.", key));
            else editor.Set(uid, id, null);
            return;
        }
        try {
            switch (kind) {
                case Kind.Integer: editor.Integer(uid, id, Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture)); break;
                case Kind.Text: editor.Set(uid, id, Text((string)value)); break;
                case Kind.Date: editor.Set(uid, id, BitConverter.GetBytes(ProjectNativeCreation.Date((DateTime)value))); break;
                case Kind.Boolean: editor.Set(uid, id, new byte[] { (bool)value ? (byte)1 : (byte)0 }); break;
                case Kind.Guid: editor.Set(uid, id, ((Guid)value).ToByteArray()); break;
                case Kind.Number: editor.Set(uid, id, Number((decimal)value * scale)); break;
                case Kind.ScaledInteger: editor.Integer(uid, id, Exact((decimal)value * scale)); break;
                case Kind.Work: editor.Set(uid, id, Number(((ProjectWork)value).Minutes * 1000)); break;
                case Kind.Units: editor.Set(uid, id, Number(((ProjectUnits)value).Value * 10000)); break;
                case Kind.Duration:
                    var duration = (ProjectDuration)value;
                    editor.Integer(uid, id, Exact(Minutes(duration) * 10));
                    if (format != 0) editor.Integer(uid, format, DurationFormat(duration)); break;
            }
        } catch (Exception ex) when (ex is OverflowException || ex is ArgumentException || ex is NotSupportedException) {
            AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_VALUE_UNREPRESENTABLE", ProjectDiagnosticSeverity.Error, ex.Message, key));
        }
    }
    private byte[] Text(string value) {
        if (value.IndexOf('\0') >= 0) throw new ArgumentException("Native strings cannot contain embedded NUL characters.");
        if (((long)value.Length + 1) * 2 > _options.MaxOutputBytes) throw new ArgumentException("Native text exceeds the output byte budget.");
        return new UnicodeEncoding(false, false, true).GetBytes(value + "\0");
    }
    private static int Exact(decimal value) { if (decimal.Truncate(value) != value) throw new ArgumentException("Native integer precision would truncate this value."); return checked((int)value); }
    private static byte[] Number(decimal value) {
        double encoded = (double)value;
        if ((decimal)encoded != value) throw new ArgumentException("Native floating-point precision would change this value.");
        return BitConverter.GetBytes(encoded);
    }
    private decimal Minutes(ProjectDuration duration) => duration.Value * (duration.Unit switch {
        ProjectDurationUnit.Minute => 1, ProjectDurationUnit.Hour => 60,
        ProjectDurationUnit.Day => duration.IsElapsed ? 1440 : _document.Settings.MinutesPerDay ?? 480,
        ProjectDurationUnit.Week => duration.IsElapsed ? 10080 : _document.Settings.MinutesPerWeek ?? 2400,
        ProjectDurationUnit.Month => duration.IsElapsed ? 43200 : (_document.Settings.MinutesPerDay ?? 480) * (_document.Settings.DaysPerMonth ?? 20),
        _ => throw new ArgumentOutOfRangeException(nameof(duration)) });
    private static int DurationFormat(ProjectDuration duration) => 3 + (int)duration.Unit * 2 + (duration.IsElapsed ? 1 : 0) + (duration.IsEstimated ? 32 : 0);
    private Guid EntityGuid(ProjectEntity entity, byte kind) {
        if (entity.Guid.HasValue) return entity.Guid.Value;
        return NativeGuid(entity.Uid, kind);
    }
    private Guid NativeGuid(int uid, byte kind) => DeriveNativeGuid(_document.NativeIdentity, uid, kind);
    private Guid SyntheticProjectSummaryGuid() => DeriveNativeGuid(_document.Guid ?? _document.NativeIdentity, 0, SyntheticProjectSummaryKind);
    internal static Guid DeriveNativeGuid(Guid identity, int uid, byte kind) {
        using var hash = System.Security.Cryptography.SHA256.Create();
        // The public project GUID is editable; implicit entity identities must survive those edits.
        byte[] seed = new byte[21]; Buffer.BlockCopy(identity.ToByteArray(), 0, seed, 0, 16);
        Buffer.BlockCopy(BitConverter.GetBytes(uid), 0, seed, 16, 4); seed[20] = kind;
        byte[] digest = hash.ComputeHash(seed); var result = new byte[16]; Buffer.BlockCopy(digest, 0, result, 0, 16); return new Guid(result);
    }
    private enum Kind { Integer, Text, Date, Boolean, Guid, Number, ScaledInteger, Work, Units, Duration }
}
