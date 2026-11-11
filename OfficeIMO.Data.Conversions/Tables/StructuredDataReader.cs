using System.Collections;
using System.Data;
using System.Data.Common;
#if NET5_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
#endif

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Reads the rows of JSON, JSON Lines, YAML, TOML or XML data as a forward-only <see cref="DbDataReader"/>, for
/// <c>SqlBulkCopy</c>, Arrow, Excel or any ADO.NET consumer. Rows come from an array (the root, each document of a stream,
/// or the path in <see cref="DataConversionOptions.Select"/>); nested objects and arrays become path columns such as
/// <c>address.city</c> and <c>tags[0]</c>. A root object whose members are all arrays is several tables, reached with
/// <see cref="NextResult"/>. The input is read twice: once to find the columns and their types, then to stream rows.
/// </summary>
public sealed class StructuredDataReader : DbDataReader {
    private readonly DataConversionPublishing.ReaderFactory _open;
    private readonly DataConversionContext _context;
    private readonly bool _text;
    private readonly IDisposable? _owned;
    private readonly List<DataPathSegment>? _select;
    private readonly bool _multiDocument;
    private readonly List<TableLayout> _tables = new();
    private readonly Mode _mode;
    private readonly CultureInfo _culture;
    private DataEventReader? _events;
    private int _table;
    private bool _inArray;
    private bool _singleDone;
    private bool _closed;
    private bool _started;
    private DataScalarKind[] _kinds = Array.Empty<DataScalarKind>();
    private string?[] _texts = Array.Empty<string?>();
    private StringBuilder?[] _joins = Array.Empty<StringBuilder?>();
    private long _row;
    private readonly long _sampleRows;
    private bool _sampled;
    private string? _latePath;

    private enum Mode {
        Rows,
        SingleRow,
        Tables
    }

    internal StructuredDataReader(DataConversionPublishing.ReaderFactory open, long? documents, DataConversionContext context, bool text, IDisposable? owned) {
        _open = open;
        _context = context;
        _text = text;
        _owned = owned;
        _culture = context.Options.Culture ?? CultureInfo.InvariantCulture;
        _multiDocument = documents is null or > 1;
        // CSV output needs only the column set, so it reads a sample first; typed output reads every row for column types.
        _sampleRows = text && context.Options.Columns == null && context.Options.SampleRows > 0 ? context.Options.SampleRows : long.MaxValue;
        if (context.Options.Select is { } select) {
            if (!DataPath.TryParse(select, out List<DataPathSegment> segments)) throw new ArgumentException("The Select path '" + select + "' is not a valid path.");
            _select = segments;
        }

        _mode = Scan();
        foreach (TableLayout table in _tables) {
            if (context.Options.Columns is { } columns) Replace(table, columns);
            table.Complete(context);
        }

        if (_tables.Count == 0) {
            var empty = new TableLayout(string.Empty);
            empty.Complete(context);
            _tables.Add(empty);
        }

        _events = open(context);
        context.Source = _events;
        Allocate();
    }

    /// <summary>Opens a JSON, JSON Lines (<c>.jsonl</c>, <c>.ndjson</c>), YAML, TOML or XML file by its extension.</summary>
    public static StructuredDataReader Open(string path, DataConversionOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A path is required.", nameof(path));
        DataFormat format = DataFormats.FromExtension(path) is { } known && !DataFormats.IsTable(known) ? known
            : throw new NotSupportedException("Open reads .json, .jsonl, .ndjson, .yaml, .yml, .toml and .xml files.");
        var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920);
        try {
            return Create(stream, format, options, owned: stream, text: false);
        } catch {
            stream.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Opens JSON, YAML, TOML or XML data from a seekable stream, which stays open. Set <see cref="DataConversionOptions.JsonLines"/>
    /// for JSON Lines.
    /// </summary>
    public static StructuredDataReader Open(Stream stream, OfficeDocumentFormat format, DataConversionOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanSeek) throw new ArgumentException("The stream must be seekable, because the rows are read in two passes.", nameof(stream));
        DataFormat source = format switch {
            OfficeDocumentFormat.Json => options?.JsonLines == true ? DataFormat.JsonLines : DataFormat.Json,
            OfficeDocumentFormat.Yaml => DataFormat.Yaml,
            OfficeDocumentFormat.Xml => DataFormat.Xml,
            OfficeDocumentFormat.Toml => DataFormat.Toml,
            _ => throw new NotSupportedException("Open reads JSON, JSON Lines, YAML, TOML and XML data.")
        };

        return Create(stream, source, options, owned: null, text: false);
    }

    internal static StructuredDataReader Create(Stream stream, DataFormat format, DataConversionOptions? options, IDisposable? owned, bool text) {
        DataConversionOptions effective = options ?? DataConversionOptions.Default;
        var context = new DataConversionContext(effective);
        long start = stream.Position;
        long? documents = format == DataFormat.JsonLines ? null : 1;
        if (format == DataFormat.Yaml) {
            documents = YamlDocumentCounter.Count(stream);
            stream.Position = start;
        }

        DataEventReader Open(DataConversionContext c) {
            stream.Position = start;
            return DataConversionEngine.OpenReader(stream, format, c);
        }

        return new StructuredDataReader(Open, documents, context, text, owned);
    }

    /// <summary>Findings so far: what the table shape could not carry, aggregated per code.</summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics => _context.CreateDiagnostics("Table");

    /// <summary>The name of the current table: the member name when the root holds several tables, otherwise empty.</summary>
    public string TableName => Current.Name;

    /// <summary>The number of tables (results) the input holds.</summary>
    public int TableCount => _tables.Count;

    internal bool IsPositional => Current.Positional;

    internal DataConversionContext Context => _context;

    private TableLayout Current => _tables[_table];

    // ---------------------------------------------------------------- pass one

    private Mode Scan() {
        // A throwaway context keeps the source's own findings from being counted twice.
        var scanContext = new DataConversionContext(_context.Options, _context.CancellationToken);
        using DataEventReader reader = _open(scanContext);
        var rows = new TableLayout(string.Empty);
        bool any = false;
        while (true) {
            DataTokenType type = TableEventWalker.Next(reader);
            if (type == DataTokenType.None) break;
            if (type != DataTokenType.StartDocument) continue;
            TableEventWalker.Next(reader);
            if (!TableEventWalker.Navigate(reader, _select)) {
                TableEventWalker.SkipToEndOfDocument(reader);
                continue;
            }

            any = true;
            if (reader.TokenType == DataTokenType.StartArray) {
                while (TableEventWalker.Next(reader) is not (DataTokenType.EndArray or DataTokenType.None)) {
                    Walk(reader, rows.Root, scan: true);
                    if (++rows.Rows >= _sampleRows) return Sampled(rows);
                }
            } else if (!_multiDocument && reader.TokenType == DataTokenType.StartObject) {
                if (ScanTables(reader)) return Mode.Tables;
                _tables.Clear();
                ScanSingleRow(rows);
                _tables.Add(rows);
                return Mode.SingleRow;
            } else {
                Walk(reader, rows.Root, scan: true);
                if (++rows.Rows >= _sampleRows && _multiDocument) return Sampled(rows);
                if (!_multiDocument) {
                    _tables.Add(rows);
                    return Mode.SingleRow;
                }
            }

            TableEventWalker.SkipToEndOfDocument(reader);
        }

        if (!any && _select != null) throw new InvalidDataException("The Select path '" + _context.Options.Select + "' matches nothing in the input.");
        _tables.Add(rows);
        return Mode.Rows;
    }

    private Mode Sampled(TableLayout rows) {
        _sampled = true;
        _tables.Add(rows);
        return Mode.Rows;
    }

    // An object whose every member is an array is a set of tables named by its keys.
    private bool ScanTables(DataEventReader reader) {
        while (TableEventWalker.Next(reader) == DataTokenType.PropertyName) {
            var table = new TableLayout(reader.Text.ToString());
            if (TableEventWalker.Next(reader) != DataTokenType.StartArray) return false;
            while (TableEventWalker.Next(reader) is not (DataTokenType.EndArray or DataTokenType.None)) {
                if (table.Rows >= _sampleRows) {
                    _sampled = true;
                    TableEventWalker.SkipValue(reader);
                    continue;
                }

                Walk(reader, table.Root, scan: true);
                table.Rows++;
            }

            _tables.Add(table);
        }

        return _tables.Count > 0;
    }

    private void ScanSingleRow(TableLayout rows) {
        var scanContext = new DataConversionContext(_context.Options, _context.CancellationToken);
        using DataEventReader reader = _open(scanContext);
        while (TableEventWalker.Next(reader) is not (DataTokenType.StartDocument or DataTokenType.None)) {
        }

        TableEventWalker.Next(reader);
        TableEventWalker.Navigate(reader, _select);
        Walk(reader, rows.Root, scan: true);
        rows.Rows = 1;
    }

    private static void Replace(TableLayout table, IReadOnlyList<string> columns) {
        var replacement = new TableLayout(table.Name);
        foreach (string column in columns) {
            if (!DataPath.TryParse(column, out List<DataPathSegment> segments)) throw new ArgumentException("The column '" + column + "' is not a valid path.");
            TableColumnNode node = replacement.Root;
            foreach (DataPathSegment segment in segments) {
                node = segment.IsIndex ? node.GetOrAddIndex(segment.Index) : node.GetOrAddName(segment.Name.AsSpan());
            }

            node.HasScalar = true;
            TableColumnNode? evidence = Find(table.Root, segments);
            if (evidence != null) {
                node.Kind = evidence.Kind;
                node.MixedKinds = evidence.MixedKinds;
                node.HasNull = evidence.HasNull;
                node.HasEmptyString = evidence.HasEmptyString;
                node.IntegerOverflow = evidence.IntegerOverflow;
                node.FloatNotDecimal = evidence.FloatNotDecimal;
            }
        }

        table.Root.Merge(replacement.Root);
    }

    private static TableColumnNode? Find(TableColumnNode root, List<DataPathSegment> segments) {
        TableColumnNode? node = root;
        foreach (DataPathSegment segment in segments) {
            node = segment.IsIndex ? node?.FindIndex(segment.Index) : node?.FindName(segment.Name.AsSpan());
        }

        return node;
    }

    // ---------------------------------------------------------------- shared walk

    /// <summary>Walks one value. Pass one grows the trie and records evidence; pass two fills cells.</summary>
    private void Walk(DataEventReader reader, TableColumnNode? node, bool scan) {
        switch (reader.TokenType) {
            case DataTokenType.Scalar:
                if (scan) node!.Observe(reader.ScalarKind, reader.Text);
                else if (node is { Column: >= 0 }) SetCell(node, reader.ScalarKind, reader.Text);
                else if (_sampled && reader.ScalarKind != DataScalarKind.Null) LateColumn(node, reader);
                return;
            case DataTokenType.StartObject: {
                bool empty = true;
                while (TableEventWalker.Next(reader) == DataTokenType.PropertyName) {
                    empty = false;
                    TableColumnNode? child = scan ? node!.GetOrAddName(reader.Text) : node?.FindName(reader.Text);
                    if (child == null && node != null && _sampled) _latePath = (node.Path.Length == 0 ? string.Empty : node.Path + ".") + reader.Text.ToString();
                    TableEventWalker.Next(reader);
                    Walk(reader, child, scan);
                }

                if (empty && scan) node!.HasEmptyCollection = true;
                return;
            }
            case DataTokenType.StartArray: {
                if (scan) node!.HasArray = true;
                if (!scan && node is { JoinsArray: true }) {
                    Join(reader, node);
                    return;
                }

                int index = 0;
                while (TableEventWalker.Next(reader) is not (DataTokenType.EndArray or DataTokenType.None)) {
                    if (scan && reader.TokenType != DataTokenType.Scalar) node!.HasNonScalarItem = true;
                    TableColumnNode? child = scan ? node!.GetOrAddIndex(index) : node?.FindIndex(index);
                    Walk(reader, child, scan);
                    index++;
                }

                if (index == 0 && scan) node!.HasEmptyCollection = true;
                return;
            }
        }
    }

    private void Join(DataEventReader reader, TableColumnNode node) {
        StringBuilder builder = _joins[node.Column] ??= new StringBuilder();
        builder.Clear();
        bool first = true;
        while (TableEventWalker.Next(reader) is not (DataTokenType.EndArray or DataTokenType.None)) {
            if (!first) builder.Append(_context.Options.ArraySeparator);
            first = false;
            if (reader.ScalarKind != DataScalarKind.Null) builder.Append(reader.Text.ToString());
        }

        _context.Report(OfficeDiagnosticCodes.ArrayJoined, OfficeConversionLossKind.Approximation,
            "Arrays were joined into one cell each; the items cannot be told apart if they contain the separator.");
        _kinds[node.Column] = DataScalarKind.String;
        _texts[node.Column] = builder.ToString();
    }

    // A row past the sample holds a value at a path the header does not have.
    private void LateColumn(TableColumnNode? node, DataEventReader reader) {
        string? path = node?.Path is { Length: > 0 } known ? known : _latePath;
        string where = path != null ? "'" + path + "'" : "a new path";
        throw new InvalidDataException("Row " + (_row + 1).ToString(CultureInfo.InvariantCulture) + " has a value at " + where
            + " that the first " + _sampleRows.ToString(CultureInfo.InvariantCulture) + " rows did not have" + (reader.Location is { } at ? " (" + at + ")" : string.Empty)
            + ". Set SampleRows to 0 to read every row before writing, or list the columns in Columns.");
    }

    private void SetCell(TableColumnNode node, DataScalarKind kind, ReadOnlySpan<char> text) {
        int column = node.Column;
        if (_text && kind == DataScalarKind.Null && node.Kind is null or DataScalarKind.String) {
            _context.Report(OfficeDiagnosticCodes.NullAsEmpty, OfficeConversionLossKind.Approximation,
                "Nulls in text columns became empty fields, which read back as empty strings.", node.Path);
        }

        _kinds[column] = kind;
        _texts[column] = kind == DataScalarKind.Null ? null : text.ToString();
        if (_text && kind == DataScalarKind.String && TableTypes.LooksTyped(text)) {
            _context.Report(OfficeDiagnosticCodes.ScalarTypeDropped, OfficeConversionLossKind.Approximation,
                "Text values that look like numbers, booleans or dates read back as those types from CSV.");
        }
    }

    private void Allocate() {
        int count = Current.Columns.Count;
        _kinds = new DataScalarKind[count];
        _texts = new string?[count];
        _joins = new StringBuilder?[count];
    }

    // ---------------------------------------------------------------- pass two

    /// <inheritdoc />
    public override bool Read() {
        if (_closed) throw new InvalidOperationException("The reader is closed.");
        _context.CancellationToken.ThrowIfCancellationRequested();
        DataEventReader reader = _events!;
        if (!_started) {
            _started = true;
            if (_mode == Mode.Tables && !EnterTables(reader)) return false;
        }

        Array.Clear(_texts, 0, _texts.Length);
        Array.Clear(_kinds, 0, _kinds.Length);
        switch (_mode) {
            case Mode.Tables:
                if (!_inArray) return false;
                if (TableEventWalker.Next(reader) is DataTokenType.EndArray or DataTokenType.None) {
                    _inArray = false;
                    return false;
                }

                return Fill(reader);
            case Mode.SingleRow:
                if (_singleDone) return false;
                _singleDone = true;
                while (TableEventWalker.Next(reader) is not (DataTokenType.StartDocument or DataTokenType.None)) {
                }

                TableEventWalker.Next(reader);
                TableEventWalker.Navigate(reader, _select);
                return Fill(reader);
            default:
                while (true) {
                    if (_inArray) {
                        if (TableEventWalker.Next(reader) is DataTokenType.EndArray or DataTokenType.None) {
                            _inArray = false;
                            TableEventWalker.SkipToEndOfDocument(reader);
                            continue;
                        }

                        return Fill(reader);
                    }

                    DataTokenType type = TableEventWalker.Next(reader);
                    if (type == DataTokenType.None) return false;
                    if (type != DataTokenType.StartDocument) continue;
                    TableEventWalker.Next(reader);
                    if (!TableEventWalker.Navigate(reader, _select)) {
                        TableEventWalker.SkipToEndOfDocument(reader);
                        continue;
                    }

                    if (reader.TokenType == DataTokenType.StartArray) {
                        _inArray = true;
                        continue;
                    }

                    bool filled = Fill(reader);
                    TableEventWalker.SkipToEndOfDocument(reader);
                    return filled;
                }
        }
    }

    private bool Fill(DataEventReader reader) {
        Walk(reader, Current.Root, scan: false);
        _row++;
        _context.ItemsWritten++;
        return true;
    }

    private bool EnterTables(DataEventReader reader) {
        while (TableEventWalker.Next(reader) is not (DataTokenType.StartDocument or DataTokenType.None)) {
        }

        TableEventWalker.Next(reader);
        TableEventWalker.Navigate(reader, _select);
        return EnterNextTable(reader);
    }

    private bool EnterNextTable(DataEventReader reader) {
        if (TableEventWalker.Next(reader) != DataTokenType.PropertyName) return false;
        TableEventWalker.Next(reader);
        _inArray = true;
        return true;
    }

    /// <inheritdoc />
    public override bool NextResult() {
        if (_mode != Mode.Tables || _table + 1 >= _tables.Count) return false;
        DataEventReader reader = _events!;
        if (!_started) {
            _started = true;
            EnterTables(reader);
        }

        while (_inArray) {
            if (TableEventWalker.Next(reader) is DataTokenType.EndArray or DataTokenType.None) _inArray = false;
            else TableEventWalker.SkipValue(reader);
        }

        _table++;
        Allocate();
        return EnterNextTable(reader);
    }

    // ---------------------------------------------------------------- values

    /// <inheritdoc />
    public override int FieldCount => Current.Columns.Count;

    /// <inheritdoc />
    public override string GetName(int ordinal) => Current.Headers[ordinal];

    /// <inheritdoc />
    public override int GetOrdinal(string name) {
        int index = Current.Headers.IndexOf(name);
        if (index < 0) throw new IndexOutOfRangeException("No column is named '" + name + "'.");
        return index;
    }

    /// <inheritdoc />
#if NET5_0_OR_GREATER
    [UnconditionalSuppressMessage("Trimming", "IL2063", Justification = "Column types are scalar value types and string, returned as data; nothing reflects over their members.")]
    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
#endif
    public override Type GetFieldType(int ordinal) => _text ? typeof(string) : Current.Types[ordinal];

    /// <inheritdoc />
    public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

    /// <inheritdoc />
    public override object GetValue(int ordinal) {
        string? text = _texts[ordinal];
        if (text == null) return DBNull.Value;
        DataScalarKind kind = _kinds[ordinal];
        return _text ? TableValues.Text(kind, text, _culture) : TableValues.Convert(kind, text, Current.Types[ordinal]);
    }

    /// <inheritdoc />
    public override int GetValues(object[] values) {
        int count = Math.Min(values.Length, FieldCount);
        for (int i = 0; i < count; i++) values[i] = GetValue(i);
        return count;
    }

    /// <inheritdoc />
    public override bool IsDBNull(int ordinal) => _texts[ordinal] == null;

    /// <inheritdoc />
    public override object this[int ordinal] => GetValue(ordinal);

    /// <inheritdoc />
    public override object this[string name] => GetValue(GetOrdinal(name));

    /// <inheritdoc />
    public override string GetString(int ordinal) => _texts[ordinal] is { } text
        ? (_text ? TableValues.Text(_kinds[ordinal], text, _culture) : text)
        : throw new InvalidCastException("The value is null.");

    /// <inheritdoc />
    public override bool GetBoolean(int ordinal) => System.Convert.ToBoolean(GetValue(ordinal), CultureInfo.InvariantCulture);

    /// <inheritdoc />
    public override byte GetByte(int ordinal) => System.Convert.ToByte(GetValue(ordinal), CultureInfo.InvariantCulture);

    /// <inheritdoc />
    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
        throw new NotSupportedException("Structured data has no binary columns.");

    /// <inheritdoc />
    public override char GetChar(int ordinal) => GetString(ordinal)[0];

    /// <inheritdoc />
    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) {
        string text = GetString(ordinal);
        if (buffer == null) return text.Length;
        int count = (int)Math.Max(0, Math.Min(length, text.Length - dataOffset));
        text.CopyTo((int)dataOffset, buffer, bufferOffset, count);
        return count;
    }

    /// <inheritdoc />
    public override DateTime GetDateTime(int ordinal) => GetValue(ordinal) switch {
        DateTime date => date,
        DateTimeOffset offset => offset.DateTime,
        object value => System.Convert.ToDateTime(value, CultureInfo.InvariantCulture)
    };

    /// <inheritdoc />
    public override decimal GetDecimal(int ordinal) => System.Convert.ToDecimal(GetValue(ordinal), CultureInfo.InvariantCulture);

    /// <inheritdoc />
    public override double GetDouble(int ordinal) => System.Convert.ToDouble(GetValue(ordinal), CultureInfo.InvariantCulture);

    /// <inheritdoc />
    public override float GetFloat(int ordinal) => System.Convert.ToSingle(GetValue(ordinal), CultureInfo.InvariantCulture);

    /// <inheritdoc />
    public override Guid GetGuid(int ordinal) => Guid.Parse(GetString(ordinal));

    /// <inheritdoc />
    public override short GetInt16(int ordinal) => System.Convert.ToInt16(GetValue(ordinal), CultureInfo.InvariantCulture);

    /// <inheritdoc />
    public override int GetInt32(int ordinal) => System.Convert.ToInt32(GetValue(ordinal), CultureInfo.InvariantCulture);

    /// <inheritdoc />
    public override long GetInt64(int ordinal) => System.Convert.ToInt64(GetValue(ordinal), CultureInfo.InvariantCulture);

    /// <inheritdoc />
#if NET5_0_OR_GREATER
    [UnconditionalSuppressMessage("Trimming", "IL2111", Justification = "The schema table stores Type values as data and does not reflect over Type members.")]
#endif
    public override DataTable GetSchemaTable() {
        var schema = new DataTable("SchemaTable") { Locale = CultureInfo.InvariantCulture };
        schema.Columns.Add(SchemaTableColumn.ColumnName, typeof(string));
        schema.Columns.Add(SchemaTableColumn.ColumnOrdinal, typeof(int));
        schema.Columns.Add(SchemaTableColumn.ColumnSize, typeof(int));
        schema.Columns.Add(SchemaTableColumn.NumericPrecision, typeof(short));
        schema.Columns.Add(SchemaTableColumn.NumericScale, typeof(short));
        schema.Columns.Add(SchemaTableColumn.DataType, typeof(Type));
        schema.Columns.Add(SchemaTableColumn.ProviderType, typeof(int));
        schema.Columns.Add(SchemaTableColumn.IsLong, typeof(bool));
        schema.Columns.Add(SchemaTableColumn.AllowDBNull, typeof(bool));
        schema.Columns.Add(SchemaTableOptionalColumn.IsReadOnly, typeof(bool));
        schema.Columns.Add(SchemaTableOptionalColumn.IsRowVersion, typeof(bool));
        schema.Columns.Add(SchemaTableColumn.IsUnique, typeof(bool));
        schema.Columns.Add(SchemaTableColumn.IsKey, typeof(bool));
        schema.Columns.Add(SchemaTableOptionalColumn.IsAutoIncrement, typeof(bool));
        schema.Columns.Add(SchemaTableColumn.BaseSchemaName, typeof(string));
        schema.Columns.Add(SchemaTableOptionalColumn.BaseCatalogName, typeof(string));
        schema.Columns.Add(SchemaTableColumn.BaseTableName, typeof(string));
        schema.Columns.Add(SchemaTableColumn.BaseColumnName, typeof(string));
        for (int i = 0; i < FieldCount; i++) {
            DataRow row = schema.NewRow();
            row[SchemaTableColumn.ColumnName] = GetName(i);
            row[SchemaTableColumn.ColumnOrdinal] = i;
            row[SchemaTableColumn.ColumnSize] = -1;
            row[SchemaTableColumn.NumericPrecision] = DBNull.Value;
            row[SchemaTableColumn.NumericScale] = DBNull.Value;
            row[SchemaTableColumn.DataType] = GetFieldType(i);
            row[SchemaTableColumn.ProviderType] = 0;
            row[SchemaTableColumn.IsLong] = false;
            row[SchemaTableColumn.AllowDBNull] = true;
            row[SchemaTableOptionalColumn.IsReadOnly] = true;
            row[SchemaTableOptionalColumn.IsRowVersion] = false;
            row[SchemaTableColumn.IsUnique] = false;
            row[SchemaTableColumn.IsKey] = false;
            row[SchemaTableOptionalColumn.IsAutoIncrement] = false;
            row[SchemaTableColumn.BaseSchemaName] = DBNull.Value;
            row[SchemaTableOptionalColumn.BaseCatalogName] = DBNull.Value;
            row[SchemaTableColumn.BaseTableName] = TableName.Length == 0 ? DBNull.Value : TableName;
            row[SchemaTableColumn.BaseColumnName] = GetName(i);
            schema.Rows.Add(row);
        }

        return schema;
    }

    /// <inheritdoc />
    public override IEnumerator GetEnumerator() => new DbEnumerator(this, closeReader: false);

    /// <inheritdoc />
    public override int Depth => 0;

    /// <inheritdoc />
    public override bool HasRows => Current.Rows > 0;

    /// <inheritdoc />
    public override bool IsClosed => _closed;

    /// <inheritdoc />
    public override int RecordsAffected => -1;

    /// <inheritdoc />
    public override void Close() {
        if (_closed) return;
        _closed = true;
        _events?.Dispose();
        _events = null;
        _owned?.Dispose();
    }

    /// <inheritdoc />
    protected override void Dispose(bool disposing) {
        if (disposing) Close();
        base.Dispose(disposing);
    }
}
