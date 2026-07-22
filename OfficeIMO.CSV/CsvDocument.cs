#nullable enable

using OfficeIMO.Drawing.Internal;
using System.Collections;
using System.Data;
using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

/// <summary>
/// Represents a CSV document with a fluent, document-centric API. Thread-safe for independent read enumeration; not thread-safe for concurrent mutations on the same instance.
/// </summary>
public sealed partial class CsvDocument
{
    private const int FileBufferSize = 256 * 1024;
    private readonly List<string> _header = new();
    private List<CsvRow> _rows = new();
    private CsvLoadMode _mode;
    private CsvStreamingSource? _streamingSource;
    private char _delimiter;
    private CultureInfo _culture;
    private Encoding _encoding;
    private CsvColumnCountMismatchPolicy _columnCountMismatchPolicy;
    private string[]? _dateTimeFormats;
    private CsvSchema? _schema;
    private bool _rowsAreParsedStringsOnly;

    /// <summary>
    /// Initializes a new in-memory CSV document with default settings.
    /// </summary>
    public CsvDocument()
    {
        _mode = CsvLoadMode.InMemory;
        _delimiter = ',';
        _culture = CultureInfo.InvariantCulture;
        _encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        _columnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict;
    }

    private CsvDocument(CsvLoadMode mode, char delimiter, CultureInfo culture, Encoding encoding, CsvColumnCountMismatchPolicy columnCountMismatchPolicy, string[]? dateTimeFormats = null)
    {
        _mode = mode;
        _delimiter = delimiter;
        _culture = culture;
        _encoding = encoding;
        _columnCountMismatchPolicy = columnCountMismatchPolicy;
        _dateTimeFormats = dateTimeFormats;
    }

    /// <summary>
    /// Creates a CSV document from a sequence of objects by projecting their properties or dictionary keys into columns.
    /// </summary>
    /// <param name="items">Sequence of objects to convert into CSV rows.</param>
    /// <param name="delimiter">Delimiter to use for the CSV document.</param>
    /// <param name="culture">Optional culture for value formatting.</param>
    /// <param name="encoding">Optional encoding for file operations.</param>
    /// <returns>A populated <see cref="CsvDocument"/>.</returns>
    [RequiresUnreferencedCode("This compatibility overload discovers properties from runtime objects. For NativeAOT, write typed rows with explicit columns or add values directly.")]
    public static CsvDocument FromObjects(IEnumerable<object?> items, char delimiter = ',', CultureInfo? culture = null, Encoding? encoding = null)
    {
        if (items == null)
        {
            throw new ArgumentNullException(nameof(items));
        }

        var list = items.ToList();
        if (list.Count == 0)
        {
            throw new ArgumentException("Provide at least one data row.", nameof(items));
        }

        var first = list.FirstOrDefault();
        if (first == null)
        {
            throw new ArgumentException("Data rows cannot be null.", nameof(items));
        }

        var columns = ObjectDataHelpers.GetColumnNames(first);
        if (columns.Count == 0)
        {
            throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
        }

        var document = new CsvDocument().WithDelimiter(delimiter);
        if (culture != null)
        {
            document.WithCulture(culture);
        }

        if (encoding != null)
        {
            document.WithEncoding(encoding);
        }

        document.WithHeader(columns.ToArray());

        foreach (var item in list)
        {
            if (item == null)
            {
                throw new InvalidOperationException("Data rows cannot contain null entries.");
            }

            var rowValues = new object?[columns.Count];
            for (var i = 0; i < columns.Count; i++)
            {
                rowValues[i] = ObjectDataHelpers.GetValue(item, columns[i]);
            }

            document.AddRow(rowValues);
        }

        return document;
    }

    /// <summary>
    /// Writes a sequence of objects directly as CSV without materializing a <see cref="CsvDocument"/>.
    /// </summary>
    /// <param name="writer">Destination text writer.</param>
    /// <param name="items">Sequence of objects to convert into CSV rows.</param>
    /// <param name="options">Optional save settings.</param>
    [RequiresUnreferencedCode("This compatibility overload discovers properties from runtime objects. For NativeAOT, write typed rows with explicit columns or add values directly.")]
    public static void WriteObjects(TextWriter writer, IEnumerable<object?> items, CsvSaveOptions? options = null)
    {
        if (writer == null)
        {
            throw new ArgumentNullException(nameof(writer));
        }

        if (items == null)
        {
            throw new ArgumentNullException(nameof(items));
        }

        options ??= new CsvSaveOptions();

        using var objectWriter = new CsvObjectWriter(writer, options, leaveOpen: true);
        var wroteAny = false;
        foreach (var item in items)
        {
            objectWriter.WriteObject(item);
            wroteAny = true;
        }

        if (!wroteAny)
        {
            throw new ArgumentException("Provide at least one data row.", nameof(items));
        }
    }

    /// <summary>
    /// Writes an <see cref="IDataReader"/> directly as CSV without materializing a <see cref="CsvDocument"/>.
    /// </summary>
    /// <param name="writer">Destination text writer.</param>
    /// <param name="reader">Source data reader positioned before the first row.</param>
    /// <param name="options">Optional save settings.</param>
    public static void WriteDataReader(TextWriter writer, IDataReader reader, CsvSaveOptions? options = null)
    {
        if (writer == null)
        {
            throw new ArgumentNullException(nameof(writer));
        }

        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        using var objectWriter = new CsvObjectWriter(writer, options ?? new CsvSaveOptions(), leaveOpen: true);
        objectWriter.WriteDataReader(reader);
    }

    /// <summary>
    /// Saves a sequence of objects directly as CSV without materializing a <see cref="CsvDocument"/>.
    /// </summary>
    /// <param name="path">Destination CSV path.</param>
    /// <param name="items">Sequence of objects to convert into CSV rows.</param>
    /// <param name="options">Optional save settings.</param>
    [RequiresUnreferencedCode("This compatibility overload discovers properties from runtime objects. For NativeAOT, write typed rows with explicit columns or add values directly.")]
    public static void SaveObjects(string path, IEnumerable<object?> items, CsvSaveOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        if (items == null)
        {
            throw new ArgumentNullException(nameof(items));
        }

        options ??= new CsvSaveOptions();
        var fullPath = Path.GetFullPath(path);
        if (options.NoClobber && File.Exists(fullPath))
        {
            throw new IOException($"The file '{fullPath}' already exists.");
        }

        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (options.Append)
        {
            using var appendWriter = CsvFile.CreateTextWriter(fullPath, options, append: true, bufferSize: FileBufferSize);
            WriteObjects(appendWriter, items, options);
            return;
        }

        var temporaryPath = Path.Combine(
            string.IsNullOrEmpty(directory) ? Environment.CurrentDirectory : directory,
            "." + Path.GetFileName(fullPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");

        try
        {
            using (var writer = CsvFile.CreateTextWriterForCompressionPath(temporaryPath, fullPath, options, bufferSize: 256 * 1024))
            {
                WriteObjects(writer, items, options);
            }

            if (File.Exists(fullPath))
            {
                File.Replace(temporaryPath, fullPath, destinationBackupFileName: null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(temporaryPath, fullPath);
            }
        }
        catch
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }

            throw;
        }
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
    /// Gets the configured date/time formats used by typed conversions.
    /// </summary>
    public IReadOnlyList<string>? DateTimeFormats => _dateTimeFormats;

    /// <summary>
    /// Gets the load mode of the document.
    /// </summary>
    public CsvLoadMode Mode => _mode;

    /// <summary>
    /// Saves the document to the specified path.
    /// </summary>
    public void Save(string path, CsvSaveOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options = ResolveSaveOptions(options);

        var fullPath = Path.GetFullPath(path);
        if (options.NoClobber && File.Exists(fullPath))
        {
            throw new IOException($"The file '{fullPath}' already exists.");
        }

        if (options.Append)
        {
            CsvCompressionType compressionType = CsvFile.ResolveCompression(options.CompressionType, fullPath);
            if (compressionType != CsvCompressionType.None)
                throw new NotSupportedException("Appending to compressed CSV files is not supported.");
            OfficeFileCommit.EnsureTargetDirectory(fullPath);
            using var writer = CsvFile.CreateTextWriter(fullPath, options, append: true, bufferSize: FileBufferSize);
            CsvWriter.Write(writer, this, options);
            return;
        }

        OfficeFileCommit.EnsureTargetDirectory(fullPath);
        var temporaryPath = OfficeFileCommit.CreateTemporaryPath(fullPath);
        try
        {
            using (var writer = CsvFile.CreateTextWriterForCompressionPath(
                       temporaryPath, fullPath, options, FileBufferSize))
            {
                CsvWriter.Write(writer, this, options);
            }
            OfficeFileCommit.CommitTemporaryFile(temporaryPath, fullPath,
                options.NoClobber ? OfficeFileCommit.ConflictPolicy.FailIfExists : OfficeFileCommit.ConflictPolicy.Replace);
            temporaryPath = string.Empty;
        }
        finally
        {
            OfficeFileCommit.DeleteIfExists(temporaryPath);
        }
    }

    /// <summary>Saves the document to a caller-owned writable stream.</summary>
    public void Save(Stream destination, CsvSaveOptions? options = null)
    {
        OfficeStreamWriter.WriteAllBytes(destination, ToBytes(options));
    }

    /// <summary>Asynchronously saves the document to a path.</summary>
    public async Task SaveAsync(string path, CsvSaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        options = ResolveSaveOptions(options);
        cancellationToken.ThrowIfCancellationRequested();
        string fullPath = Path.GetFullPath(path);
        if (options.NoClobber && File.Exists(fullPath)) throw new IOException($"The file '{fullPath}' already exists.");
        CsvCompressionType compressionType = CsvFile.ResolveCompression(options.CompressionType, fullPath);

        if (options.Append)
        {
            if (compressionType != CsvCompressionType.None)
                throw new NotSupportedException("Appending to compressed CSV files is not supported.");
            OfficeFileCommit.EnsureTargetDirectory(fullPath);
            byte[] appendBytes = SerializeToBytes(options, CsvCompressionType.None);
            using var stream = new FileStream(fullPath, FileMode.Append, FileAccess.Write, FileShare.Read,
                FileBufferSize, FileOptions.Asynchronous);
            int appendOffset = GetAppendOffset(appendBytes, options.Encoding, stream.Length > 0);
#if NET6_0_OR_GREATER
            await stream.WriteAsync(appendBytes.AsMemory(appendOffset), cancellationToken).ConfigureAwait(false);
#else
            await stream.WriteAsync(appendBytes, appendOffset, appendBytes.Length - appendOffset, cancellationToken).ConfigureAwait(false);
#endif
            await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
            return;
        }

        byte[] bytes = SerializeToBytes(options, compressionType);
        await OfficeFileCommit.WriteAllBytesAsync(fullPath, bytes,
            options.NoClobber ? OfficeFileCommit.ConflictPolicy.FailIfExists : OfficeFileCommit.ConflictPolicy.Replace,
            cancellationToken).ConfigureAwait(false);
    }

    private static int GetAppendOffset(byte[] bytes, Encoding? configuredEncoding, bool destinationHasContent)
    {
        if (!destinationHasContent) return 0;
        Encoding encoding = configuredEncoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        byte[] preamble = encoding.GetPreamble();
        if (preamble.Length == 0 || bytes.Length < preamble.Length) return 0;
        for (int index = 0; index < preamble.Length; index++)
        {
            if (bytes[index] != preamble[index]) return 0;
        }

        return preamble.Length;
    }

    /// <summary>Asynchronously saves the document to a caller-owned writable stream.</summary>
    public Task SaveAsync(Stream destination, CsvSaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeStreamWriter.WriteAllBytesAsync(destination, ToBytes(options), cancellationToken);
    }

    /// <summary>Encodes the document using the selected CSV encoding and compression.</summary>
    public byte[] ToBytes(CsvSaveOptions? options = null)
    {
        options = ResolveSaveOptions(options);
        if (options.Append || options.NoClobber)
            throw new ArgumentException("Append and NoClobber apply only to path saves.", nameof(options));
        CsvCompressionType compressionType = options.CompressionType == CsvCompressionType.Auto
            ? CsvCompressionType.None
            : options.CompressionType;
        return SerializeToBytes(options, compressionType);
    }

    /// <summary>Encodes the document in a new writable memory stream positioned at the beginning.</summary>
    public MemoryStream ToStream(CsvSaveOptions? options = null) => new MemoryStream(ToBytes(options));

    private byte[] SerializeToBytes(CsvSaveOptions options, CsvCompressionType compressionType)
    {
        var serializationOptions = CopySaveOptions(options, compressionType);
        using var stream = new MemoryStream();
        using (TextWriter writer = CsvFile.CreateTextWriter(stream, serializationOptions, leaveOpen: true, FileBufferSize))
        {
            CsvWriter.Write(writer, this, serializationOptions);
        }
        return stream.ToArray();
    }

    private static CsvSaveOptions CopySaveOptions(CsvSaveOptions source, CsvCompressionType compressionType)
    {
        return new CsvSaveOptions
        {
            Delimiter = source.Delimiter,
            DelimiterText = source.DelimiterText,
            NewLine = source.NewLine,
            IncludeHeader = source.IncludeHeader,
            Culture = source.Culture,
            Encoding = source.Encoding,
            CompressionType = compressionType,
            CompressionLevel = source.CompressionLevel,
            NullValue = source.NullValue,
            DateTimeFormat = source.DateTimeFormat,
            UseUtc = source.UseUtc,
            FormulaInjectionPolicy = source.FormulaInjectionPolicy,
            QuoteMode = source.QuoteMode,
            QuoteFields = source.QuoteFields
        };
    }

    private CsvSaveOptions ResolveSaveOptions(CsvSaveOptions? options)
    {
        return options ?? new CsvSaveOptions
        {
            Delimiter = _delimiter,
            Culture = _culture,
            Encoding = _encoding
        };
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
                var aligned = AlignParsedObjectValues(values, _header.Count, _columnCountMismatchPolicy);
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
    /// Sets additional date/time formats used by typed row conversion and schema validation.
    /// </summary>
    public CsvDocument WithDateTimeFormats(params string[] formats)
    {
        if (formats == null)
        {
            throw new ArgumentNullException(nameof(formats));
        }

        _dateTimeFormats = formats.ToArray();
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
        _rowsAreParsedStringsOnly = false;
        var aligned = NormalizeValuesLength(values);
        _rows.Add(new CsvRow(this, aligned));
    }

    private void AddParsedRowInternal(IReadOnlyList<string> values, CsvLoadOptions options)
    {
        _rowsAreParsedStringsOnly = (_rows.Count == 0 || _rowsAreParsedStringsOnly) &&
            options.NullValue is null &&
            (options.StaticColumns is null || options.StaticColumns.Count == 0);
        _rows.Add(new CsvRow(this, BuildParsedObjectValues(values, _header.Count, options)));
    }

    private void SetHeader(IEnumerable<string> headers)
    {
        _header.Clear();
        _header.AddRange(headers);
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
