using System.Data.Common;
using OfficeIMO.CSV;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Reads rows from any <see cref="DbDataReader"/> (CSV, Excel, a database) as data events: an array of objects, or an
/// object of arrays named by table when there are several results. Path-shaped headers (<c>address.city</c>,
/// <c>tags[0]</c>) rebuild nested data. Untyped text is inferred per column in a first pass, so a column never mixes
/// numbers and strings.
/// </summary>
internal sealed class TableEventReader : DataEventReader {
    private readonly Func<DbDataReader> _open;
    private readonly bool _infer;
    private readonly bool _positional;
    private readonly IReadOnlyList<string>? _tableNames;
    private readonly char _decimalSeparator;
    private readonly List<DataScalarKind?[]> _inferred = new();
    private readonly Queue<Pending> _pending = new();
    private DbDataReader? _reader;
    private int _result;
    private long _row;
    private Phase _phase;
    private PlanNode? _plan;
    private DataScalarKind?[]? _kinds;

    private enum Phase {
        Start,
        Rows,
        Done
    }

    private readonly struct Pending {
        public Pending(DataTokenType type, DataScalarKind kind = default, string? text = null) {
            Type = type;
            Kind = kind;
            Text = text;
        }

        public DataTokenType Type { get; }
        public DataScalarKind Kind { get; }
        public string? Text { get; }
    }

    private sealed class PlanNode {
        public string? Name;
        public int Column = -1;
        public bool IsArray;
        public readonly List<PlanNode> Children = new();
    }

    /// <param name="open">Opens the source; called twice when inferring types.</param>
    /// <param name="infer">Whether string cells are typed by column inference.</param>
    /// <param name="positional">Whether the rows have no header and become arrays.</param>
    /// <param name="tableNames">The result names, when the source holds several tables that become an object of arrays.</param>
    /// <param name="context">The conversion context.</param>
    public TableEventReader(Func<DbDataReader> open, bool infer, bool positional, IReadOnlyList<string>? tableNames, DataConversionContext context) : base(context) {
        _open = open;
        _infer = infer;
        _positional = positional;
        _tableNames = tableNames is { Count: > 1 } ? tableNames : null;
        _decimalSeparator = (context.Options.Culture ?? CultureInfo.InvariantCulture).NumberFormat.NumberDecimalSeparator[0];
        if (infer) InferColumns();
        _reader = open();
    }

    /// <summary>Opens CSV with the strict profile, or TSV with the TSV profile (or the caller's reader options), and per-column inference.</summary>
    public static TableEventReader OpenCsv(Stream input, DataConversionContext context, bool tsv) {
        CsvReaderOptions options = context.Options.CsvReader?.Clone() ?? CsvProfiles.CreateReaderOptions(tsv ? CsvProfile.Tsv : CsvProfile.Strict);
        if (context.Options.Culture != null) options.Culture = context.Options.Culture;
        bool infer = context.Options.InferTypes ?? true;
        if (infer && !input.CanSeek) throw new ArgumentException("Inferring CSV column types needs a seekable stream.", nameof(input));
        long start = input.CanSeek ? input.Position : 0;
        DbDataReader Open() {
            if (input.CanSeek) input.Position = start;
            return CsvDocument.OpenReader(input, options);
        }

        return new TableEventReader(Open, infer, !options.HasHeaderRow, null, context);
    }

    public override string? Location => _reader == null ? null : "row " + _row.ToString(CultureInfo.InvariantCulture);

    private void InferColumns() {
        using DbDataReader reader = _open();
        do {
            var kinds = new DataScalarKind?[reader.FieldCount];
            var mixed = new bool[reader.FieldCount];
            while (reader.Read()) {
                Context.CancellationToken.ThrowIfCancellationRequested();
                for (int i = 0; i < kinds.Length; i++) {
                    if (mixed[i] || reader.IsDBNull(i)) continue;
                    object value = reader.GetValue(i);
                    if (value is not string text) continue;
                    DataScalarKind? kind = TableTypes.Infer(text.AsSpan(), _decimalSeparator);
                    if (kind == null) continue;
                    if (kinds[i] is not { } seen) kinds[i] = kind;
                    else if (seen != kind) {
                        if ((seen == DataScalarKind.Integer && kind == DataScalarKind.Float) || (seen == DataScalarKind.Float && kind == DataScalarKind.Integer)) {
                            kinds[i] = DataScalarKind.Float;
                        } else {
                            kinds[i] = DataScalarKind.String;
                            mixed[i] = true;
                        }
                    }
                }
            }

            _inferred.Add(kinds);
        } while (reader.NextResult());
    }

    protected override bool ReadCore() {
        while (_pending.Count == 0) {
            if (!Advance()) return false;
        }

        Pending next = _pending.Dequeue();
        switch (next.Type) {
            case DataTokenType.PropertyName: SetPropertyName(next.Text.AsSpan()); break;
            case DataTokenType.Scalar: SetScalar(next.Kind, next.Text.AsSpan()); break;
            default: Set(next.Type); break;
        }

        return true;
    }

    private bool Advance() {
        DbDataReader reader = _reader!;
        switch (_phase) {
            case Phase.Start:
                _pending.Enqueue(new Pending(DataTokenType.StartDocument));
                if (_tableNames != null) _pending.Enqueue(new Pending(DataTokenType.StartObject));
                BeginResult(reader);
                _phase = Phase.Rows;
                return true;
            case Phase.Rows:
                Context.CancellationToken.ThrowIfCancellationRequested();
                if (reader.Read()) {
                    _row++;
                    Emit(_plan!, reader);
                    return true;
                }

                _pending.Enqueue(new Pending(DataTokenType.EndArray));
                if (_tableNames != null && reader.NextResult()) {
                    _result++;
                    BeginResult(reader);
                    return true;
                }

                if (_tableNames != null) _pending.Enqueue(new Pending(DataTokenType.EndObject));
                _pending.Enqueue(new Pending(DataTokenType.EndDocument));
                _phase = Phase.Done;
                return true;
            default:
                return false;
        }
    }

    private void BeginResult(DbDataReader reader) {
        if (_tableNames != null) _pending.Enqueue(new Pending(DataTokenType.PropertyName, text: _result < _tableNames.Count ? _tableNames[_result] : "table" + _result.ToString(CultureInfo.InvariantCulture)));
        _pending.Enqueue(new Pending(DataTokenType.StartArray));
        _kinds = _infer && _result < _inferred.Count ? _inferred[_result] : null;
        _row = 0;
        _plan = BuildPlan(reader);
    }

    private PlanNode BuildPlan(DbDataReader reader) {
        int count = reader.FieldCount;
        var root = new PlanNode { IsArray = _positional };
        if (_positional) {
            for (int i = 0; i < count; i++) root.Children.Add(new PlanNode { Column = i });
            return root;
        }

        if (Context.Options.NestedHeaders && TryNested(reader, root)) return root;
        root.Children.Clear();
        for (int i = 0; i < count; i++) root.Children.Add(new PlanNode { Name = reader.GetName(i), Column = i });
        return root;
    }

    // Headers rebuild nesting only when every header is a path and no path is both a value and a container.
    private bool TryNested(DbDataReader reader, PlanNode root) {
        bool nested = false;
        for (int i = 0; i < reader.FieldCount; i++) {
            string header = reader.GetName(i);
            if (!DataPath.TryParse(header, out List<DataPathSegment> segments) || segments[0].IsIndex) return false;
            if (segments.Count > 1) nested = true;
            PlanNode node = root;
            for (int s = 0; s < segments.Count; s++) {
                DataPathSegment segment = segments[s];
                if (node.Column >= 0 || (node.Children.Count > 0 && node.IsArray != segment.IsIndex)) return Conflict(header);
                node.IsArray = segment.IsIndex;
                PlanNode? child = null;
                if (segment.IsIndex) {
                    while (node.Children.Count <= segment.Index) node.Children.Add(new PlanNode());
                    child = node.Children[segment.Index];
                } else {
                    foreach (PlanNode existing in node.Children) {
                        if (existing.Name == segment.Name) {
                            child = existing;
                            break;
                        }
                    }

                    if (child == null) {
                        child = new PlanNode { Name = segment.Name };
                        node.Children.Add(child);
                    }
                }

                if (s == segments.Count - 1) {
                    if (child.Column >= 0 || child.Children.Count > 0) return Conflict(header);
                    child.Column = i;
                }

                node = child;
            }
        }

        return nested;
    }

    private bool Conflict(string header) {
        Context.Report(OfficeDiagnosticCodes.HeaderPathAmbiguous, OfficeConversionLossKind.Approximation,
            "Headers did not form a consistent set of paths, so they were kept as flat keys.", "column '" + header + "'");
        return false;
    }

    private void Emit(PlanNode node, DbDataReader reader) {
        if (node.Column >= 0) {
            EmitCell(reader, node.Column);
            return;
        }

        if (node.Children.Count == 0 && !node.IsArray && node.Name == null) {
            _pending.Enqueue(new Pending(DataTokenType.Scalar, DataScalarKind.Null));
            return;
        }

        if (node.IsArray) {
            int last = node.Children.Count - 1;
            while (last >= 0 && !HasValue(node.Children[last], reader)) last--;
            _pending.Enqueue(new Pending(DataTokenType.StartArray));
            for (int i = 0; i <= last; i++) Emit(node.Children[i], reader);
            _pending.Enqueue(new Pending(DataTokenType.EndArray));
            return;
        }

        _pending.Enqueue(new Pending(DataTokenType.StartObject));
        foreach (PlanNode child in node.Children) {
            _pending.Enqueue(new Pending(DataTokenType.PropertyName, text: child.Name));
            Emit(child, reader);
        }

        _pending.Enqueue(new Pending(DataTokenType.EndObject));
    }

    private bool HasValue(PlanNode node, DbDataReader reader) {
        // An empty cell is an absent item: CSV cannot tell a missing trailing item from an empty one.
        if (node.Column >= 0) return !reader.IsDBNull(node.Column) && reader.GetValue(node.Column) is not string { Length: 0 };
        foreach (PlanNode child in node.Children) {
            if (HasValue(child, reader)) return true;
        }

        return false;
    }

    private DataScalarKind? Kind(int column) => _kinds != null && column < _kinds.Length ? _kinds[column] : null;

    private void EmitCell(DbDataReader reader, int column) {
        if (reader.IsDBNull(column)) {
            _pending.Enqueue(new Pending(DataTokenType.Scalar, DataScalarKind.Null));
            return;
        }

        object value = reader.GetValue(column);
        if (value is string text) {
            DataScalarKind kind = _infer ? Kind(column) ?? DataScalarKind.String : DataScalarKind.String;
            if (text.Length == 0) {
                _pending.Enqueue(kind == DataScalarKind.String ? new Pending(DataTokenType.Scalar, DataScalarKind.String, string.Empty)
                    : new Pending(DataTokenType.Scalar, DataScalarKind.Null));
                return;
            }

            string normalized = kind switch {
                DataScalarKind.Integer or DataScalarKind.Float => TableTypes.NumberText(text, _decimalSeparator),
                DataScalarKind.Boolean => text.Length == 4 ? "true" : "false",
                _ => text
            };
            _pending.Enqueue(new Pending(DataTokenType.Scalar, kind, normalized));
            return;
        }

        (DataScalarKind typed, string typedText) = Typed(value);
        _pending.Enqueue(new Pending(DataTokenType.Scalar, typed, typedText));
    }

    internal static (DataScalarKind Kind, string Text) Typed(object value) {
        switch (value) {
            case bool boolean:
                return (DataScalarKind.Boolean, boolean ? "true" : "false");
            case long or int or short or byte or sbyte or ushort or uint or ulong:
                return (DataScalarKind.Integer, System.Convert.ToString(value, CultureInfo.InvariantCulture)!);
            case decimal number:
                string decimalText = number.ToString(CultureInfo.InvariantCulture);
                return (decimalText.IndexOf('.') < 0 ? DataScalarKind.Integer : DataScalarKind.Float, decimalText);
            case double or float: {
                double number = System.Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (double.IsNaN(number)) return (DataScalarKind.Float, "NaN");
                if (double.IsInfinity(number)) return (DataScalarKind.Float, number > 0 ? "Infinity" : "-Infinity");
                if (Math.Floor(number) == number && Math.Abs(number) < 9007199254740992d) {
                    return (DataScalarKind.Integer, ((long)number).ToString(CultureInfo.InvariantCulture));
                }

                return (DataScalarKind.Float, number.ToString("R", CultureInfo.InvariantCulture));
            }
            case DateTimeOffset offset:
                return (DataScalarKind.OffsetDateTime, offset.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz", CultureInfo.InvariantCulture));
            case DateTime date:
                return date.TimeOfDay == TimeSpan.Zero
                    ? (DataScalarKind.LocalDate, date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    : (DataScalarKind.LocalDateTime, date.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture));
            case TimeSpan time:
                return (DataScalarKind.LocalTime, time.ToString("hh\\:mm\\:ss\\.FFFFFFF", CultureInfo.InvariantCulture).TrimEnd('.'));
            default:
                return (DataScalarKind.String, System.Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
        }
    }

    public override void Dispose() {
        _reader?.Dispose();
        _reader = null;
    }
}
