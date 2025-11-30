#nullable enable

using System.Globalization;
using System.Text;

namespace OfficeIMO.CSV;

/// <summary>
/// Represents a CSV document with a fluent, document-centric API. Thread-safe for independent read enumeration; not thread-safe for concurrent mutations on the same instance.
/// </summary>
public sealed class CsvDocument
{
    private readonly List<string> _header = new();
    private List<CsvRow> _rows = new();
    private CsvLoadMode _mode;
    private CsvStreamingSource? _streamingSource;
    private char _delimiter;
    private CultureInfo _culture;
    private Encoding _encoding;
    private CsvSchema? _schema;

    /// <summary>
    /// Initializes a new in-memory CSV document with default settings.
    /// </summary>
    public CsvDocument()
    {
        _mode = CsvLoadMode.InMemory;
        _delimiter = ',';
        _culture = CultureInfo.InvariantCulture;
        _encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
    }

    private CsvDocument(CsvLoadMode mode, char delimiter, CultureInfo culture, Encoding encoding)
    {
        _mode = mode;
        _delimiter = delimiter;
        _culture = culture;
        _encoding = encoding;
    }

    /// <summary>
    /// Gets the document header columns.
    /// </summary>
    public IReadOnlyList<string> Header => _header;

    /// <summary>
    /// Gets the delimiter used by the document.
    /// </summary>
    public char Delimiter => _delimiter;

    /// <summary>
    /// Gets the culture used for type conversions.
    /// </summary>
    public CultureInfo Culture => _culture;

    /// <summary>
    /// Gets the encoding used when reading or writing files.
    /// </summary>
    public Encoding Encoding => _encoding;

    /// <summary>
    /// Gets the load mode of the document.
    /// </summary>
    public CsvLoadMode Mode => _mode;

    /// <summary>
    /// Loads a CSV document from disk.
    /// </summary>
    public static CsvDocument Load(string path, CsvLoadOptions? options = null)
    {
        options ??= new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        return LoadInternal(() => new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true), options, encoding);
    }

    /// <summary>
    /// Parses a CSV document from text.
    /// </summary>
    public static CsvDocument Parse(string text, CsvLoadOptions? options = null)
    {
        options ??= new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        return LoadInternal(() => new StringReader(text), options, encoding);
    }

    private static CsvDocument LoadInternal(Func<TextReader> readerFactory, CsvLoadOptions options, Encoding encoding)
    {
        var document = new CsvDocument(options.Mode, options.Delimiter, options.Culture, encoding);

        using var reader = readerFactory();
        using var enumerator = CsvParser.Parse(reader, options).GetEnumerator();

        if (!enumerator.MoveNext())
        {
            return document;
        }

        var firstRecord = enumerator.Current;

        if (options.HasHeaderRow)
        {
            document.SetHeader(firstRecord);
        }
        else
        {
            document.SetHeader(GenerateDefaultHeader(firstRecord.Length));
            document.AddRowInternal(firstRecord.Cast<object?>().ToArray());
        }

        if (options.Mode == CsvLoadMode.InMemory)
        {
            while (enumerator.MoveNext())
            {
                document.AddRowInternal(enumerator.Current.Cast<object?>().ToArray());
            }
        }
        else
        {
            document._streamingSource = new CsvStreamingSource(readerFactory, options, skipFirstRecord: options.HasHeaderRow);
        }

        return document;
    }

    /// <summary>
    /// Saves the document to the specified path.
    /// </summary>
    public CsvDocument Save(string path, CsvSaveOptions? options = null)
    {
        options ??= new CsvSaveOptions
        {
            Delimiter = _delimiter,
            Culture = _culture,
            Encoding = _encoding
        };

        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        using var writer = new StreamWriter(path, append: false, encoding);
        CsvWriter.Write(writer, this, options);
        return this;
    }

    /// <summary>
    /// Serializes the document to a string using the provided save options.
    /// </summary>
    public string ToString(CsvSaveOptions? options)
    {
        options ??= new CsvSaveOptions
        {
            Delimiter = _delimiter,
            Culture = _culture,
            Encoding = _encoding
        };

        using var writer = new StringWriter();
        CsvWriter.Write(writer, this, options);
        return writer.ToString();
    }

    /// <inheritdoc />
    public override string ToString() => ToString(null);

    /// <summary>
    /// Returns rows as an enumerable sequence. In streaming mode this is lazy.
    /// </summary>
    public IEnumerable<CsvRow> AsEnumerable()
    {
        if (_mode == CsvLoadMode.InMemory)
        {
            return _rows;
        }

        if (_streamingSource is null)
        {
            return _rows;
        }

        return StreamRows();

        IEnumerable<CsvRow> StreamRows()
        {
            foreach (var values in _streamingSource.ReadRows())
            {
                var aligned = NormalizeValuesLength(values);
                yield return new CsvRow(this, aligned);
            }
        }
    }

    /// <summary>
    /// Sets the delimiter used for reading and writing.
    /// </summary>
    public CsvDocument WithDelimiter(char delimiter)
    {
        _delimiter = delimiter;
        return this;
    }

    /// <summary>
    /// Replaces the header row.
    /// </summary>
    public CsvDocument WithHeader(params string[] headers)
    {
        if (headers is null || headers.Length == 0)
        {
            throw new ArgumentException("Header must contain at least one column.", nameof(headers));
        }

        EnsureInMemory();
        ValidateExistingRows(headers.Length);

        _header.Clear();
        _header.AddRange(headers);
        return this;
    }

    /// <summary>
    /// Sets the culture for type conversions.
    /// </summary>
    public CsvDocument WithCulture(CultureInfo culture)
    {
        _culture = culture ?? throw new ArgumentNullException(nameof(culture));
        return this;
    }

    /// <summary>
    /// Sets the encoding used for file operations.
    /// </summary>
    public CsvDocument WithEncoding(Encoding encoding)
    {
        _encoding = encoding ?? throw new ArgumentNullException(nameof(encoding));
        return this;
    }

    /// <summary>
    /// Adds a new row to the document.
    /// </summary>
    public CsvDocument AddRow(params object?[] values)
    {
        EnsureInMemory();
        AddRowInternal(values);
        return this;
    }

    /// <summary>
    /// Adds an existing row instance to the document.
    /// </summary>
    public CsvDocument AddRow(CsvRow row)
    {
        EnsureInMemory();
        if (row is null)
        {
            throw new ArgumentNullException(nameof(row));
        }

        if (row.FieldCount != _header.Count)
        {
            throw new CsvException($"Row contains {row.FieldCount} fields but header defines {_header.Count} columns.");
        }

        _rows.Add(row.CloneFor(this));
        return this;
    }

    /// <summary>
    /// Adds a computed column to the document.
    /// </summary>
    public CsvDocument AddColumn(string name, Func<CsvRow, object?> valueFactory)
    {
        EnsureInMemory();
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Column name cannot be null or empty.", nameof(name));
        }

        if (valueFactory is null)
        {
            throw new ArgumentNullException(nameof(valueFactory));
        }

        if (_header.Any(h => string.Equals(h, name, StringComparison.OrdinalIgnoreCase)))
        {
            throw new CsvException($"Column '{name}' already exists.");
        }

        _header.Add(name);
        for (var i = 0; i < _rows.Count; i++)
        {
            var value = valueFactory(_rows[i]);
            _rows[i].Values = AppendValue(_rows[i].Values, value);
        }

        return this;
    }

    /// <summary>
    /// Removes a column by name.
    /// </summary>
    public CsvDocument RemoveColumn(string name)
    {
        EnsureInMemory();
        if (!TryGetColumnIndexInternal(name, out var index))
        {
            return this;
        }

        _header.RemoveAt(index);
        foreach (var row in _rows)
        {
            var newValues = new object?[row.Values.Length - 1];
            if (index > 0)
            {
                Array.Copy(row.Values, 0, newValues, 0, index);
            }

            var remaining = row.Values.Length - index - 1;
            if (remaining > 0)
            {
                Array.Copy(row.Values, index + 1, newValues, index, remaining);
            }

            row.Values = newValues;
        }

        return this;
    }

    /// <summary>
    /// Filters rows using the provided predicate.
    /// </summary>
    public CsvDocument Filter(Func<CsvRow, bool> predicate)
    {
        EnsureInMemory();
        if (predicate is null)
        {
            throw new ArgumentNullException(nameof(predicate));
        }

        _rows = _rows.Where(predicate).ToList();
        return this;
    }

    /// <summary>
    /// Sorts rows by a column name.
    /// </summary>
    public CsvDocument SortBy(string columnName, bool descending = false, IComparer<object?>? comparer = null)
    {
        EnsureInMemory();
        var index = GetColumnIndex(columnName);
        comparer ??= NullAwareComparer;
        _rows = descending
            ? _rows.OrderByDescending(r => r.Values[index], comparer).ToList()
            : _rows.OrderBy(r => r.Values[index], comparer).ToList();
        return this;
    }

    /// <summary>
    /// Sorts rows using a typed key selector.
    /// </summary>
    public CsvDocument SortBy<TKey>(Func<CsvRow, TKey> keySelector, bool descending = false, IComparer<TKey>? comparer = null)
    {
        EnsureInMemory();
        _rows = descending
            ? _rows.OrderByDescending(keySelector, comparer ?? Comparer<TKey>.Default).ToList()
            : _rows.OrderBy(keySelector, comparer ?? Comparer<TKey>.Default).ToList();
        return this;
    }

    /// <summary>
    /// Executes a custom transformation.
    /// </summary>
    public CsvDocument Transform(Func<CsvDocument, CsvDocument> transformer)
    {
        if (transformer is null)
        {
            throw new ArgumentNullException(nameof(transformer));
        }

        return transformer(this);
    }

    /// <summary>
    /// Attaches a schema to the document.
    /// </summary>
    public CsvDocument EnsureSchema(Action<CsvSchemaBuilder> buildAction)
    {
        if (buildAction is null)
        {
            throw new ArgumentNullException(nameof(buildAction));
        }

        var builder = new CsvSchemaBuilder();
        buildAction(builder);
        _schema = builder.Build();
        return this;
    }

    /// <summary>
    /// Validates the document against the configured schema.
    /// </summary>
    public CsvDocument Validate(out IReadOnlyList<CsvValidationError> errors)
    {
        if (_schema is null)
        {
            errors = Array.Empty<CsvValidationError>();
            return this;
        }

        errors = CsvValidator.Validate(this, _schema);
        return this;
    }

    /// <summary>
    /// Validates and throws when validation fails.
    /// </summary>
    public CsvDocument ValidateOrThrow()
    {
        Validate(out var errors);
        if (errors.Count > 0)
        {
            throw new CsvValidationException("CSV document failed validation.", errors);
        }

        return this;
    }

    /// <summary>
    /// Forces a streaming document to materialize into memory, enabling transformations.
    /// </summary>
    public CsvDocument Materialize()
    {
        if (_mode == CsvLoadMode.InMemory)
        {
            return this;
        }

        _rows = AsEnumerable().Select(r => r.CloneFor(this)).ToList();
        _streamingSource = null;
        _mode = CsvLoadMode.InMemory;
        return this;
    }

    internal int GetColumnIndex(string columnName)
    {
        if (!TryGetColumnIndexInternal(columnName, out var index))
        {
            throw new CsvException($"Column '{columnName}' was not found in the header.");
        }

        return index;
    }

    internal bool TryGetColumnIndexInternal(string columnName, out int index)
    {
        index = _header.FindIndex(h => string.Equals(h, columnName, StringComparison.OrdinalIgnoreCase));
        return index >= 0;
    }

    private void AddRowInternal(object?[] values)
    {
        var aligned = NormalizeValuesLength(values);
        _rows.Add(new CsvRow(this, aligned));
    }

    private void SetHeader(IEnumerable<string> headers)
    {
        _header.Clear();
        _header.AddRange(headers);
    }

    private object?[] NormalizeValuesLength(IEnumerable<object?> values)
    {
        var array = values.ToArray();
        if (_header.Count == 0)
        {
            _header.AddRange(GenerateDefaultHeader(array.Length));
            return array;
        }

        if (array.Length != _header.Count)
        {
            throw new CsvException($"Row contains {array.Length} values but header defines {_header.Count} columns.");
        }

        return array;
    }

    private void ValidateExistingRows(int expectedColumns)
    {
        if (_rows.Any(r => r.FieldCount != expectedColumns))
        {
            throw new CsvException("Existing rows do not match the new header length.");
        }
    }

    private void EnsureInMemory()
    {
        if (_mode == CsvLoadMode.Stream)
        {
            throw new InvalidOperationException("Operation requires in-memory mode. Call Materialize() or load with CsvLoadMode.InMemory.");
        }
    }

    private static IReadOnlyList<string> GenerateDefaultHeader(int count)
    {
        var result = new List<string>(count);
        for (var i = 0; i < count; i++)
        {
            result.Add($"Column{i + 1}");
        }

        return result;
    }

    private static object?[] AppendValue(IReadOnlyList<object?> source, object? value)
    {
        var result = new object?[source.Count + 1];
        for (var i = 0; i < source.Count; i++)
        {
            result[i] = source[i];
        }

        result[result.Length - 1] = value;
        return result;
    }

    private static readonly IComparer<object?> NullAwareComparer = Comparer<object?>.Create((x, y) =>
    {
        if (ReferenceEquals(x, y))
        {
            return 0;
        }

        if (x is null)
        {
            return -1;
        }

        if (y is null)
        {
            return 1;
        }

        if (x is IComparable comparable)
        {
            return comparable.CompareTo(y);
        }

        var xs = x.ToString() ?? string.Empty;
        var ys = y.ToString() ?? string.Empty;
        return string.CompareOrdinal(xs, ys);
    });
}
