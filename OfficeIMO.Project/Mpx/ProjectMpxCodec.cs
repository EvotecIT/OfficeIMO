namespace OfficeIMO.Project;

/// <summary>Immutable source bytes and the omissions that matter when rewriting or converting MPX.</summary>
internal sealed class ProjectMpxSource {
    internal byte[] Bytes = Array.Empty<byte>();
    internal long ModelRevision;
    internal int CodePage = 1252;
    internal char Separator = ',';
    internal string CurrencyPosition = "1", DateFormat = "0", BarDateFormat = "0";
    internal IReadOnlyList<string[]> Comments = Array.Empty<string[]>();
    internal IReadOnlyList<ProjectDiagnostic> Unmodeled = Array.Empty<ProjectDiagnostic>();
    internal IReadOnlyList<ProjectDiagnostic> Unrepresented = Array.Empty<ProjectDiagnostic>();
}

internal static partial class ProjectMpxCodec {
    internal static ProjectDocument Read(byte[] bytes, ProjectLoadOptions options, CancellationToken token) {
        var records = ProjectMpxRecords.Read(bytes, options, token);
        var reader = new Reader(options, token);
        try {
            reader.Read(records);
            reader.Document.MpxSource = new ProjectMpxSource { Bytes = bytes, ModelRevision = reader.Document.Revision, CodePage = records.CodePage, Separator = records.Separator,
                CurrencyPosition = reader.CurrencyPosition, DateFormat = reader.DateFormat, BarDateFormat = reader.BarDateFormat, Comments = reader.Comments.ToArray(), Unmodeled = reader.Unmodeled.ToArray() };
            reader.Document.FinishRead(options);
            return reader.Document;
        } catch { reader.Document.DisposeFailedRead(); throw; }
    }

    private sealed partial class Reader {
        internal readonly ProjectDocument Document = ProjectDocument.CreateForRead();
        internal readonly List<ProjectDiagnostic> Unmodeled = new List<ProjectDiagnostic>();
        internal readonly List<string[]> Comments = new List<string[]>();
        internal string CurrencyPosition = "1", DateFormat = "0", BarDateFormat = "0";
        private readonly ProjectLoadOptions _options;
        private readonly CancellationToken _token;
        private readonly ProjectMpxValues _values = new ProjectMpxValues();
        private readonly Dictionary<string, ProjectCalendar> _calendars = new Dictionary<string, ProjectCalendar>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<int, ProjectTask> _taskRows = new Dictionary<int, ProjectTask>();
        private readonly Dictionary<int, ProjectResource> _resourceRows = new Dictionary<int, ProjectResource>();
        private readonly List<(ProjectTask Task, string Text, bool Uids, bool Successors)> _links = new List<(ProjectTask, string, bool, bool)>();
        private readonly Stack<ProjectTask> _parents = new Stack<ProjectTask>();
        private ProjectTask? _task;
        private ProjectResource? _resource;
        private ProjectCalendar? _calendar;
        private int[]? _taskFields, _resourceFields;
        private int _row;
        private int _phase;
        private char _listSeparator = ',';
        private string _projectCalendar = "Standard";
        private readonly HashSet<int> _singletons = new HashSet<int>();
        internal Reader(ProjectLoadOptions options, CancellationToken token) {
            _options = options; _token = token; Document.ReadDiagnosticLimit = options.MaxDiagnostics;
            _values.OnLoss = message => Opaque(message);
        }
        private static string Get(string[] record, int index) => ProjectMpxValues.Get(record, index);
        private static bool Has(string[] record, int index) => !ProjectMpxValues.Empty(Get(record, index));
        private string Location => "/MPX/Record[" + _row + "]";
        private void Opaque(string message, string? location = null) {
            var diagnostic = new ProjectDiagnostic("PROJECT_MPX_UNMODELED", ProjectDiagnosticSeverity.Warning, message, location ?? Location, true);
            Document.AddReadDiagnostic(diagnostic);
            if (Unmodeled.Count < _options.MaxDiagnostics) Unmodeled.Add(diagnostic);
        }
        private void Tail(string[] record, int first) {
            if (record.Skip(first).Any(v => !ProjectMpxValues.Empty(v))) Opaque("Additional MPX fields remain in the original bytes and have no typed mapping.");
        }
        private void Budget() {
            if ((long)Document.TaskIndex.Count + Document.ResourceIndex.Count + Document.CalendarIndex.Count + Document.AssignmentIndex.Count >= _options.MaxEntities)
                throw new InvalidDataException("MPX exceeds MaxEntities.");
        }
        private void Phase(int phase) {
            if (phase < _phase) throw new InvalidDataException("MPX records are out of order at " + Location + ".");
            _phase = phase;
        }
        internal void Read(ProjectMpxRecords input) {
            _listSeparator = input.Separator;
            for (int i = 1; i < input.Records.Count; i++) {
                _token.ThrowIfCancellationRequested(); _row = i + 1;
                var r = input.Records[i]; int kind = ProjectMpxValues.Integer(r[0]);
                if ((kind == 10 || kind == 11 || kind == 12 || kind == 30 || kind == 40 || kind == 41 || kind == 60 || kind == 61) && !_singletons.Add(kind))
                    throw new InvalidDataException("Duplicate MPX definition record " + kind + ".");
                switch (kind) {
                    case 0: Comments.Add(r); break;
                    case 10: Phase(10); Currency(r); break;
                    case 11: Phase(11); Defaults(r); break;
                    case 12: Phase(12); DateSettings(r); break;
                    case 20: Phase(20); BaseCalendar(r); break;
                    case 25: RequireCalendar(false); Hours(r); break;
                    case 26: RequireCalendar(false); Exception(r); break;
                    case 30: Phase(30); Header(r); _calendar = null; break;
                    case 40: Phase(40); _resourceFields = TextFields(r, false); break;
                    case 41: Phase(41); _resourceFields = NumericFields(r); break;
                    case 50: Phase(50); ReadResource(r); break;
                    case 51: if (_resource == null || _phase != 50) throw new InvalidDataException("MPX resource notes have no owning resource."); _resource.Notes = Get(r, 1).Replace("\u007f", "\n"); Tail(r, 2); break;
                    case 55: if (_resource == null || _phase != 50) throw new InvalidDataException("MPX resource calendar has no owning resource."); ResourceCalendar(r); break;
                    case 56: RequireCalendar(true); Hours(r); break;
                    case 57: RequireCalendar(true); Exception(r); break;
                    case 60: Phase(60); _taskFields = TextFields(r, true); _calendar = null; break;
                    case 61: Phase(61); _taskFields = NumericFields(r); _calendar = null; break;
                    case 70: Phase(70); ReadTask(r); break;
                    case 71: RequireTask(); _task!.Notes = Get(r, 1).Replace("\u007f", "\n"); Tail(r, 2); break;
                    case 72: RequireTask(); Opaque("Recurring-task metadata remains inert in the original MPX bytes.", "/Task[UID=" + _task!.Uid + "]/Recurrence"); break;
                    case 75: RequireTask(); Assignment(r); break;
                    case 76: RequireTask(); Opaque("Assignment workgroup metadata remains inert in the original MPX bytes."); break;
                    case 80: case 81: Phase(80); Opaque("External project, DDE, and OLE references remain inert; no file, application, or network resource is opened."); break;
                    default: Opaque("Unknown MPX record " + kind + " remains in the original bytes."); break;
                }
            }
            if (!_calendars.TryGetValue(_projectCalendar, out var projectCalendar)) {
                if (!string.Equals(_projectCalendar, "Standard", StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("MPX refers to an undefined project calendar: " + _projectCalendar);
                Budget(); projectCalendar = Document.Calendars.AddStandardWorkingWeek(); _calendars.Add("Standard", projectCalendar);
            }
            Document.Calendar = projectCalendar;
            ResolveLinks();
        }
        private void RequireTask() { if (_task == null || _phase != 70) throw new InvalidDataException("MPX task child record has no owning task."); }
        private void RequireCalendar(bool resource) {
            if (_calendar == null || (resource ? _phase != 50 : _phase != 20)) throw new InvalidDataException("MPX calendar detail has no owning definition.");
        }
        private void Currency(string[] r) {
            if (Has(r, 1)) Document.Settings.CurrencySymbol = _values.Currency = r[1];
            if (Has(r, 2) && (ProjectMpxValues.Integer(r[2]) < 0 || ProjectMpxValues.Integer(r[2]) > 3)) throw new InvalidDataException("Invalid MPX currency position.");
            if (Has(r, 2)) CurrencyPosition = r[2];
            if (Has(r, 3)) { int digits = ProjectMpxValues.Integer(r[3]); if (digits < 0 || digits > 2) throw new InvalidDataException("Invalid MPX currency digits."); Document.Settings.CurrencyDigits = digits; }
            if (Has(r, 4)) _values.GroupSeparator = r[4];
            if (Has(r, 5)) _values.DecimalSeparator = r[5];
            if (_values.DecimalSeparator == _values.GroupSeparator) throw new InvalidDataException("MPX decimal and group separators must differ.");
            Tail(r, 6);
        }
        private void Defaults(string[] r) {
            int Unit(int index, int fallback) { int unit = Has(r, index) ? ProjectMpxValues.Integer(r[index]) : fallback; if (unit < 0 || unit > 3) throw new InvalidDataException("Invalid MPX default unit."); return unit; }
            _values.DefaultDurationUnit = Unit(1, 2); _values.DefaultWorkUnit = Unit(3, 1);
            if (Has(r, 2)) Document.Settings.DefaultTaskType = _values.Flag(r[2]) ? ProjectTaskType.FixedDuration : ProjectTaskType.FixedUnits;
            int Minutes(int index, int fallback) { decimal minutes = Has(r, index) ? checked(_values.Number(r[index]) * 60) : fallback; if (minutes <= 0 || minutes != decimal.Truncate(minutes)) throw new InvalidDataException("MPX working hours must represent positive whole minutes."); return checked((int)minutes); }
            Document.Settings.MinutesPerDay = _values.MinutesPerDay = Minutes(4, 480);
            Document.Settings.MinutesPerWeek = _values.MinutesPerWeek = Minutes(5, 2400);
            // These settings have effects beyond their scalar values; do not pretend they were applied.
            if (r.Skip(6).Any(v => !ProjectMpxValues.Empty(v))) Opaque("Default rates, status propagation, and split settings have no typed mapping.");
        }
        private void DateSettings(string[] r) {
            if (Has(r, 1)) _values.DateOrder = ProjectMpxValues.Integer(r[1]);
            if (_values.DateOrder < 0 || _values.DateOrder > 2) throw new InvalidDataException("Invalid MPX date order.");
            if (Has(r, 2) && Get(r, 2) != "0" && Get(r, 2) != "1") throw new InvalidDataException("Invalid MPX time format.");
            if (Has(r, 3)) { int minutes = ProjectMpxValues.Integer(r[3]); if (minutes < 0 || minutes >= 1440) throw new InvalidDataException("Invalid MPX default time."); Document.Settings.DefaultStartTime = _values.DefaultTime = TimeSpan.FromMinutes(minutes); }
            if (Has(r, 4)) _values.DateSeparator = r[4]; if (Has(r, 5)) _values.TimeSeparator = r[5];
            if (Has(r, 6)) _values.Am = r[6]; if (Has(r, 7)) _values.Pm = r[7];
            if (Has(r, 8)) DateFormat = r[8]; if (Has(r, 9)) BarDateFormat = r[9];
            Tail(r, 10);
        }
        private void Header(string[] r) {
            Document.Name = Get(r, 1).Length != 0 ? r[1] : null; Document.Company = Get(r, 2).Length != 0 ? r[2] : null; Document.Manager = Get(r, 3).Length != 0 ? r[3] : null;
            if (Get(r, 4).Length != 0) _projectCalendar = r[4];
            if (Has(r, 5)) Document.Settings.StartDate = _values.Date(r[5]);
            if (Has(r, 6)) Document.Settings.FinishDate = _values.Date(r[6]);
            if (Has(r, 7)) Document.Settings.ScheduleFromStart = !_values.Flag(r[7]);
            Tail(r, 8);
        }
        private static int[] NumericFields(string[] r) {
            var values = r.Skip(1).Select(ProjectMpxValues.Integer).ToArray();
            if (values.Length < 2 || values.Any(v => v < 0) || values.Distinct().Count() != values.Length) throw new InvalidDataException("Invalid MPX numeric field definition.");
            return values;
        }
        private static int[] TextFields(string[] r, bool task) {
            var known = task ? ProjectMpxFields.Tasks.ToDictionary(f => f.Name, f => f.Id, StringComparer.OrdinalIgnoreCase) : ProjectMpxFields.Resources.ToDictionary(f => f.Name, f => f.Id, StringComparer.OrdinalIgnoreCase);
            known["Unique ID"] = task ? 98 : 49;
            if (task) { known["Outline Level"] = 3; known["Summary"] = 120; known["Predecessors"] = 70; known["Successors"] = 71; known["Unique ID Predecessors"] = 74; known["Unique ID Successors"] = 75; }
            else known["Base Calendar"] = 48;
            foreach (var custom in ProjectMpxFields.CustomMappings(task)) known[custom.Name] = custom.Id;
            known["Baseline Work"] = 21; known["Baseline Cost"] = 31;
            if (task) { known["Baseline Duration"] = 41; known["Baseline Start"] = 56; known["Baseline Finish"] = 57; }
            if (r.Length < 3) throw new InvalidDataException("An MPX table requires at least two fields.");
            var result = r.Skip(1).Select((name, i) => known.TryGetValue(name, out int id) ? id : -i - 1).ToArray();
            if (result.Distinct().Count() != result.Length) throw new InvalidDataException("Duplicate MPX text field declaration.");
            return result;
        }
    }
}
