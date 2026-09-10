using OfficeIMO.Core.Internal;
using System.Text;

namespace OfficeIMO.Project;

/// <summary>Canonical MPX output with a full typed-value inventory and pre-commit loss assessment.</summary>
internal sealed partial class ProjectMpxWriter {
    private readonly ProjectDocument _document;
    private readonly ProjectSaveOptions _options;
    private readonly CancellationToken _token;
    private readonly Dictionary<string, object?> _snapshot;
    private readonly HashSet<string> _handled = new HashSet<string>(StringComparer.Ordinal);
    private readonly List<ProjectDiagnostic> _diagnostics = new List<ProjectDiagnostic>();
    private readonly MemoryStream? _output;
    private long _length;
    private readonly char _separator;
    private readonly int _codePage;
    private readonly ProjectMpxValues _values;
    private readonly Dictionary<ProjectTask, int> _taskRows = new Dictionary<ProjectTask, int>();
    private readonly Dictionary<ProjectResource, int> _resourceRows = new Dictionary<ProjectResource, int>();
    private readonly Dictionary<ProjectCalendar, string> _calendarNames = new Dictionary<ProjectCalendar, string>();
    private ProjectMpxWriter(ProjectDocument document, ProjectSaveOptions options, bool serialize, CancellationToken token) {
        _document = document; _options = options; _token = token; _snapshot = ProjectModelSnapshot.Capture(document, token);
        _separator = options.MpxSeparator ?? document.MpxSource?.Separator ?? ',';
        _codePage = (int?)options.MpxEncoding ?? document.MpxSource?.CodePage ?? 1252;
        if (serialize) _output = new MemoryStream();
        _values = new ProjectMpxValues { MinutesPerDay = document.Settings.MinutesPerDay ?? 480, MinutesPerWeek = document.Settings.MinutesPerWeek ?? 2400 };
    }
    internal static bool CanRetain(ProjectDocument document, ProjectSaveOptions options) => options.PreserveUnchangedBytes && document.MpxSource != null
        && document.MpxSource.ModelRevision == document.Revision && !document.HasPendingBatchChanges
        && (!options.MpxEncoding.HasValue || (int)options.MpxEncoding.Value == document.MpxSource.CodePage)
        && (!options.MpxSeparator.HasValue || options.MpxSeparator.Value == document.MpxSource.Separator);
    internal static (ProjectReport Report, byte[]? Bytes) Plan(ProjectDocument document, ProjectSaveOptions options, bool serialize, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (CanRetain(document, options)) {
            var source = document.MpxSource!;
            var findings = source.Unrepresented.ToList();
            if (source.Bytes.LongLength > options.MaxOutputBytes) findings.Add(new ProjectDiagnostic("PROJECT_MPX_OUTPUT_LIMIT", ProjectDiagnosticSeverity.Error, "MPX exceeds MaxOutputBytes.", "/"));
            return (new ProjectReport(document.Revision, findings), serialize ? source.Bytes : null);
        }
        var writer = new ProjectMpxWriter(document, options, serialize, token);
        try {
            try { writer.Build(); }
            catch (Exception ex) when (ex is ArgumentException || ex is InvalidDataException || ex is OverflowException || ex is NotSupportedException || ex is FormatException) {
                writer.Diagnostic("PROJECT_MPX_UNREPRESENTABLE", ex.Message, "/", false);
            }
            var report = new ProjectReport(document.Revision, writer._diagnostics);
            return (report, serialize && !report.HasErrors && !(report.HasLoss && options.LossPolicy == OfficeConversionLossPolicy.Block) ? writer._output!.ToArray() : null);
        } finally { writer._output?.Dispose(); }
    }
    private void Diagnostic(string code, string message, string path, bool loss = true) {
        if (_diagnostics.Count < 1000) _diagnostics.Add(new ProjectDiagnostic(code, loss ? ProjectDiagnosticSeverity.Warning : ProjectDiagnosticSeverity.Error, message, path, loss));
        else if (_diagnostics.Count == 1000) _diagnostics.Add(new ProjectDiagnostic("PROJECT_MPX_DIAGNOSTICS_LIMIT", ProjectDiagnosticSeverity.Error, "MPX assessment exceeded its diagnostic budget.", "/"));
    }
    private static string Path(ProjectEntity entity, string type) => "/" + type + "[UID=" + entity.Uid + "]";
    private void Handle(string key) => _handled.Add(key);
    private void HandleTree(string path) { foreach (string key in _snapshot.Keys.Where(k => k == path || k.StartsWith(path + "/", StringComparison.Ordinal) || k.StartsWith(path + "[", StringComparison.Ordinal))) Handle(key); }
    private string Value(string path, object? value) { Handle(path); return ProjectMpxValues.Text(value); }
    private void Record(params string?[] fields) {
        _token.ThrowIfCancellationRequested();
        string record = ProjectMpxRecords.WriteRecord(fields, _separator);
        if (record.Length > _options.MaxOutputBytes - _length) throw new InvalidDataException("MPX exceeds MaxOutputBytes.");
        byte[] bytes = OfficeLegacySingleByteEncoding.Encode(record, _codePage); _length = checked(_length + bytes.Length);
        _output?.Write(bytes, 0, bytes.Length);
    }
    private void Build() {
        if (_document.TaskIndex.Count > 9999 || _document.Resources.Count > 9999 || _document.Calendars.Count > 250)
            throw new NotSupportedException("MPX supports at most 9999 tasks, 9999 resources, and 250 base calendars.");
        Record("MPX", "OfficeIMO", "4.0", _codePage switch { 1252 => "ANSI", 437 => "437", 850 => "850", _ => "MAC" });
        foreach (var comment in _document.MpxSource?.Comments ?? Array.Empty<string[]>()) Record(comment);
        Settings(); Calendars(); Header(); Resources(); Tasks();
        Handle("/Dependency/Count"); Handle("/Definition/Count");
        for (int i = 0; i < _document.CustomFields.Count; i++) {
            var definition = _document.CustomFields[i];
            var mapping = ProjectMpxFields.CustomMappings(true).Concat(ProjectMpxFields.CustomMappings(false)).FirstOrDefault(m => m.FieldId == definition.FieldId);
            if (mapping == null) continue;
            Handle("/Definition[" + i + "]/FieldId");
            if (definition.FieldName == mapping.Name) Handle("/Definition[" + i + "]/FieldName");
        }
        foreach (var pair in _snapshot) {
            _token.ThrowIfCancellationRequested();
            if (pair.Value != null && !_handled.Contains(pair.Key)) Diagnostic("PROJECT_MPX_FIELD_LOSS", "This modeled value has no MPX 4.0 representation and is omitted.", pair.Key);
            else if (pair.Value is DateTime date && date.Ticks % TimeSpan.TicksPerMinute != 0)
                Diagnostic("PROJECT_MPX_DATE_PRECISION", "Qualified MPX dates require whole minutes.", pair.Key, false);
            else if (pair.Value is string text && text != text.Trim(' ', '\t'))
                Diagnostic("PROJECT_MPX_WHITESPACE", "MPX trims leading and trailing spaces and tabs in fields.", pair.Key);
        }
        if (_document.Source?.HasOpaqueStructures == true) Diagnostic("PROJECT_MPX_XML_CONTENT_LOSS", "Unmodeled XML elements and attributes are omitted from MPX.", "/");
        if (_document.NativeSource != null) Diagnostic("PROJECT_MPX_NATIVE_CONTENT_LOSS", "Native presentation, macros, embedded content, signatures, and unmodeled records are omitted from MPX.", "/");
        foreach (var diagnostic in _document.MpxSource?.Unmodeled ?? Array.Empty<ProjectDiagnostic>())
            Diagnostic("PROJECT_MPX_SOURCE_CONTENT_LOSS", diagnostic.Message + " Rewriting omits this unmodeled content.", diagnostic.Location);
    }
    private void Settings() {
        var s = _document.Settings;
        if (s.CurrencyDigits > 2) throw new NotSupportedException("MPX supports at most two currency decimal digits.");
        Record("10", Value("/Settings/CurrencySymbol", s.CurrencySymbol ?? "$"), _document.MpxSource?.CurrencyPosition ?? "1", Value("/Settings/CurrencyDigits", s.CurrencyDigits ?? 2), ",", ".");
        Handle("/Settings/MinutesPerDay"); Handle("/Settings/MinutesPerWeek");
        Record("11", "2", s.DefaultTaskType == null ? "" : s.DefaultTaskType == ProjectTaskType.FixedDuration ? "1" : "0", "1",
            ProjectMpxValues.Text(_values.MinutesPerDay / 60m), ProjectMpxValues.Text(_values.MinutesPerWeek / 60m));
        if (s.DefaultTaskType != ProjectTaskType.FixedWork) Handle("/Settings/DefaultTaskType");
        if (s.DaysPerMonth == 20) Handle("/Settings/DaysPerMonth");
        if (s.NewTasksAreManual == false) Handle("/Settings/NewTasksAreManual");
        Handle("/Settings/DefaultStartTime");
        decimal time = (s.DefaultStartTime ?? TimeSpan.FromHours(8)).Ticks / (decimal)TimeSpan.TicksPerMinute;
        if (time != decimal.Truncate(time)) throw new NotSupportedException("MPX default time is measured in whole minutes.");
        Record("12", "2", "1", ProjectMpxValues.Text(time), "/", ":", "AM", "PM", _document.MpxSource?.DateFormat ?? "0", _document.MpxSource?.BarDateFormat ?? "0");
    }
    private void Header() {
        string calendar = "Standard";
        if (_document.Calendar != null) { calendar = _calendarNames[_document.Calendar]; Handle("/Settings/Calendar"); }
        Record("30", Value("/Project/Name", _document.Name), Value("/Project/Company", _document.Company), Value("/Project/Manager", _document.Manager), calendar,
            Value("/Settings/StartDate", _document.Settings.StartDate), Value("/Settings/FinishDate", _document.Settings.FinishDate),
            _document.Settings.ScheduleFromStart.HasValue ? _document.Settings.ScheduleFromStart.Value ? "0" : "1" : "");
        Handle("/Settings/ScheduleFromStart");
    }
}
