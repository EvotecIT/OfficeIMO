#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Read;

namespace OfficeIMO.Tabular;

/// <summary>
/// Provides one forward-only, read-only tabular API over delimited text and Excel workbooks.
/// </summary>
public sealed partial class TabularReader : DbDataReader {
    private readonly IReadOnlyList<string> _tableNames;
    private readonly Func<int, DbDataReader> _openTable;
    private readonly IDisposable? _owner;
    private readonly CultureInfo _culture;
    private DbDataReader _current;
    private int _tableIndex;
    private bool _closed;

    private TabularReader(
        TabularFormat format,
        IReadOnlyList<string> tableNames,
        Func<int, DbDataReader> openTable,
        IDisposable? owner,
        CultureInfo culture) {
        if (tableNames.Count == 0) {
            owner?.Dispose();
            throw new InvalidDataException("The tabular source does not contain any readable tables.");
        }

        Format = format;
        _tableNames = tableNames;
        _openTable = openTable;
        _owner = owner;
        _culture = culture;
        _current = openTable(0);
    }

    /// <summary>Gets the detected physical format.</summary>
    public TabularFormat Format { get; }

    /// <summary>Gets the names of the results exposed by this reader.</summary>
    public IReadOnlyList<string> TableNames => _tableNames;

    /// <summary>Gets the current result name.</summary>
    public string TableName => _tableNames[_tableIndex];

    /// <summary>Opens a file using format detection based on its extension.</summary>
    public static TabularReader Open(string path, TabularReadOptions? options = null) =>
        Open(path, TabularFormat.Auto, options);

    /// <summary>Opens a file using the requested physical format.</summary>
    public static TabularReader Open(string path, TabularFormat format, TabularReadOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        if (!File.Exists(path)) {
            throw new FileNotFoundException($"File '{path}' does not exist.", path);
        }

        var effectiveOptions = options ?? new TabularReadOptions();
        effectiveOptions.Validate();
        effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
        long inputLength = new FileInfo(path).Length;
        if (inputLength > effectiveOptions.MaxInputBytes) {
            throw new InvalidDataException(
                $"Tabular input contains {inputLength} bytes, exceeding the configured limit of {effectiveOptions.MaxInputBytes} bytes.");
        }

        TabularFormat resolvedFormat = format == TabularFormat.Auto ? DetectFormat(path) : format;
        return resolvedFormat switch {
            TabularFormat.DelimitedText => OpenDelimitedPath(path, effectiveOptions),
            TabularFormat.ExcelOpenXml => OpenExcelPath(path, effectiveOptions, TabularFormat.ExcelOpenXml),
            TabularFormat.ExcelBinary => OpenExcelBinaryPath(path, effectiveOptions),
            _ => throw new NotSupportedException($"Tabular format '{resolvedFormat}' is not supported.")
        };
    }

    /// <summary>Opens a caller-owned stream using an explicit physical format.</summary>
    public static TabularReader Open(
        Stream stream,
        TabularFormat format,
        TabularReadOptions? options = null,
        string? sourceName = null) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        if (format == TabularFormat.Auto) {
            if (string.IsNullOrWhiteSpace(sourceName)) {
                throw new ArgumentException("A source name is required when detecting a stream format.", nameof(sourceName));
            }

            format = DetectFormat(sourceName!);
        }

        var effectiveOptions = options ?? new TabularReadOptions();
        effectiveOptions.Validate();
        effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
        if (stream.CanSeek) {
            long remainingBytes = checked(stream.Length - stream.Position);
            if (remainingBytes > effectiveOptions.MaxInputBytes) {
                throw new InvalidDataException(
                    $"Tabular input contains {remainingBytes} unread bytes, exceeding the configured limit of {effectiveOptions.MaxInputBytes} bytes.");
            }
        }

        return format switch {
            TabularFormat.DelimitedText => OpenDelimitedStream(stream, sourceName, effectiveOptions),
            TabularFormat.ExcelOpenXml => OpenExcelStream(stream, effectiveOptions, TabularFormat.ExcelOpenXml),
            TabularFormat.ExcelBinary => OpenExcelBinaryStream(stream, effectiveOptions),
            _ => throw new NotSupportedException($"Tabular format '{format}' is not supported.")
        };
    }

    /// <inheritdoc />
    public override bool NextResult() {
        ThrowIfClosed();
        if (_tableIndex + 1 >= _tableNames.Count) {
            return false;
        }

        _current.Dispose();
        _tableIndex++;
        _current = _openTable(_tableIndex);
        return true;
    }

    /// <inheritdoc />
    public override void Close() => CloseCore();

    /// <inheritdoc />
    protected override void Dispose(bool disposing) {
        if (disposing) {
            CloseCore();
        }

        base.Dispose(disposing);
    }

    private void CloseCore() {
        if (_closed) {
            return;
        }

        _closed = true;
        _current.Dispose();
        _owner?.Dispose();
    }

    /// <inheritdoc />
    public override object this[int ordinal] => _current[ordinal];

    /// <inheritdoc />
    public override object this[string name] => _current[name];

    /// <inheritdoc />
    public override int Depth => _current.Depth;

    /// <inheritdoc />
    public override int FieldCount => _current.FieldCount;

    /// <inheritdoc />
    public override bool HasRows => _current.HasRows;

    /// <inheritdoc />
    public override bool IsClosed => _closed || _current.IsClosed;

    /// <inheritdoc />
    public override int RecordsAffected => _current.RecordsAffected;

    /// <inheritdoc />
    public override int VisibleFieldCount => _current.VisibleFieldCount;

    /// <inheritdoc />
    public override bool GetBoolean(int ordinal) => _current.GetBoolean(ordinal);

    /// <inheritdoc />
    public override byte GetByte(int ordinal) => _current.GetByte(ordinal);

    /// <inheritdoc />
    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
        _current.GetBytes(ordinal, dataOffset, buffer, bufferOffset, length);

    /// <inheritdoc />
    public override char GetChar(int ordinal) => _current.GetChar(ordinal);

    /// <inheritdoc />
    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) =>
        _current.GetChars(ordinal, dataOffset, buffer, bufferOffset, length);

    /// <inheritdoc />
    public override string GetDataTypeName(int ordinal) => _current.GetDataTypeName(ordinal);

    /// <inheritdoc />
    public override DateTime GetDateTime(int ordinal) => _current.GetDateTime(ordinal);

    /// <inheritdoc />
    public override decimal GetDecimal(int ordinal) => _current.GetDecimal(ordinal);

    /// <inheritdoc />
    public override double GetDouble(int ordinal) => _current.GetDouble(ordinal);

    /// <inheritdoc />
    public override IEnumerator GetEnumerator() => _current.GetEnumerator();

    /// <inheritdoc />
#if NET8_0_OR_GREATER
    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
#endif
    public override Type GetFieldType(int ordinal) => _current.GetFieldType(ordinal);

    /// <inheritdoc />
    public override T GetFieldValue<T>(int ordinal) {
        Type destinationType = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
        if (destinationType == typeof(string)) {
            return (T)(object)GetString(ordinal);
        }
        if (destinationType == typeof(bool)) {
            return (T)(object)GetBoolean(ordinal);
        }
        if (destinationType == typeof(byte)) {
            return (T)(object)GetByte(ordinal);
        }
        if (destinationType == typeof(short)) {
            return (T)(object)GetInt16(ordinal);
        }
        if (destinationType == typeof(int)) {
            return (T)(object)GetInt32(ordinal);
        }
        if (destinationType == typeof(long)) {
            return (T)(object)GetInt64(ordinal);
        }
        if (destinationType == typeof(float)) {
            return (T)(object)GetFloat(ordinal);
        }
        if (destinationType == typeof(double)) {
            return (T)(object)GetDouble(ordinal);
        }
        if (destinationType == typeof(decimal)) {
            return (T)(object)GetDecimal(ordinal);
        }
        if (destinationType == typeof(DateTime)) {
            return (T)(object)GetDateTime(ordinal);
        }
        if (destinationType == typeof(Guid)) {
            return (T)(object)GetGuid(ordinal);
        }

        object value = GetValue(ordinal);
        if (value is T typed) {
            return typed;
        }

        return (T)Convert.ChangeType(value, destinationType, _culture);
    }

    /// <inheritdoc />
    public override float GetFloat(int ordinal) => _current.GetFloat(ordinal);

    /// <inheritdoc />
    public override Guid GetGuid(int ordinal) => _current.GetGuid(ordinal);

    /// <inheritdoc />
    public override short GetInt16(int ordinal) => _current.GetInt16(ordinal);

    /// <inheritdoc />
    public override int GetInt32(int ordinal) => _current.GetInt32(ordinal);

    /// <inheritdoc />
    public override long GetInt64(int ordinal) => _current.GetInt64(ordinal);

    /// <inheritdoc />
    public override string GetName(int ordinal) => _current.GetName(ordinal);

    /// <inheritdoc />
    public override int GetOrdinal(string name) => _current.GetOrdinal(name);

    /// <inheritdoc />
    public override string GetString(int ordinal) => _current.GetString(ordinal);

    /// <inheritdoc />
    public override object GetValue(int ordinal) => _current.GetValue(ordinal);

    /// <inheritdoc />
    public override int GetValues(object[] values) => _current.GetValues(values);

    /// <inheritdoc />
    public override bool IsDBNull(int ordinal) => _current.IsDBNull(ordinal);

    /// <inheritdoc />
    public override bool Read() {
        ThrowIfClosed();
        return _current.Read();
    }

    /// <inheritdoc />
    public override DataTable? GetSchemaTable() => _current.GetSchemaTable();

    private void ThrowIfClosed() {
        if (IsClosed) {
            throw new InvalidOperationException("The tabular reader is closed.");
        }
    }

    private static TabularFormat DetectFormat(string sourceName) {
        string extension = Path.GetExtension(sourceName);
        if (string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".tsv", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".txt", StringComparison.OrdinalIgnoreCase)) {
            return TabularFormat.DelimitedText;
        }

        if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".xltx", StringComparison.OrdinalIgnoreCase)
            || string.Equals(extension, ".xltm", StringComparison.OrdinalIgnoreCase)) {
            return TabularFormat.ExcelOpenXml;
        }

        if (string.Equals(extension, ".xlsb", StringComparison.OrdinalIgnoreCase)) {
            return TabularFormat.ExcelBinary;
        }

        throw new NotSupportedException($"Cannot detect a tabular format from '{sourceName}'.");
    }

    private static TabularReader OpenDelimitedPath(string path, TabularReadOptions options) {
        CsvLoadOptions loadOptions = CreateCsvOptions(path, options);
        var readerOptions = new CsvDataReaderOptions {
            InferSchema = options.InferTypes,
            SchemaSampleSize = options.SchemaSampleRows
        };
        CsvDataReader reader = CsvDocument.CreateDataReader(path, loadOptions, readerOptions);
        string tableName = options.TableName ?? Path.GetFileNameWithoutExtension(path);
        return new TabularReader(
            TabularFormat.DelimitedText,
            new[] { tableName },
            _ => reader,
            owner: null,
            options.Culture);
    }

    private static TabularReader OpenDelimitedStream(Stream stream, string? sourceName, TabularReadOptions options) {
        CsvLoadOptions loadOptions = CreateCsvOptions(sourceName, options);
        var readerOptions = new CsvDataReaderOptions {
            InferSchema = options.InferTypes,
            SchemaSampleSize = options.SchemaSampleRows
        };
        CsvDataReader reader = CsvDocument.CreateDataReader(stream, loadOptions, readerOptions);
        string tableName = options.TableName
            ?? (string.IsNullOrWhiteSpace(sourceName) ? "Data" : Path.GetFileNameWithoutExtension(sourceName));
        return new TabularReader(
            TabularFormat.DelimitedText,
            new[] { tableName },
            _ => reader,
            owner: null,
            options.Culture);
    }

    private static CsvLoadOptions CreateCsvOptions(string? sourceName, TabularReadOptions options) {
        char delimiter = options.Delimiter
            ?? (string.Equals(Path.GetExtension(sourceName), ".tsv", StringComparison.OrdinalIgnoreCase) ? '\t' : ',');
        return new CsvLoadOptions {
            Mode = CsvLoadMode.Stream,
            HasHeaderRow = options.HasHeaderRow,
            Delimiter = delimiter,
            DetectDelimiter = options.DetectDelimiter,
            TrimWhitespace = options.TrimWhitespace,
            Encoding = options.Encoding,
            Culture = options.Culture,
            MaxInputBytes = options.MaxInputBytes,
            MaxDecompressedBytes = options.MaxInputBytes,
            CancellationToken = options.CancellationToken
        };
    }

    private static TabularReader OpenExcelPath(
        string path,
        TabularReadOptions options,
        TabularFormat format) {
        ExcelDocumentReader owner = ExcelDocumentReader.Open(path, CreateExcelOptions(options));
        return CreateExcelReader(owner, options, format);
    }

    private static TabularReader OpenExcelStream(
        Stream stream,
        TabularReadOptions options,
        TabularFormat format) {
        ExcelDocumentReader owner = ExcelDocumentReader.Open(stream, CreateExcelOptions(options));
        return CreateExcelReader(owner, options, format);
    }

    private static TabularReader CreateExcelReader(
        ExcelDocumentReader owner,
        TabularReadOptions options,
        TabularFormat format) {
        IReadOnlyList<string> workbookTables = owner.GetSheetNames();
        IReadOnlyList<string> selectedTables = SelectTables(workbookTables, options.TableName);
        try {
            return new TabularReader(
                format,
                selectedTables,
                tableIndex => {
                    ExcelSheetReader sheet = owner.GetSheet(selectedTables[tableIndex]);
                    IDataReader reader = sheet.ReadUsedRangeAsDataReader(
                        headersInFirstRow: options.HasHeaderRow,
                        schemaSampleRows: options.InferTypes ? options.SchemaSampleRows : 0,
                        ct: options.CancellationToken);
                    return (DbDataReader)reader;
                },
                owner,
                options.Culture);
        } catch {
            owner.Dispose();
            throw;
        }
    }

    private static IReadOnlyList<string> SelectTables(IReadOnlyList<string> tableNames, string? selectedTable) {
        if (string.IsNullOrWhiteSpace(selectedTable)) {
            return tableNames;
        }

        string? match = tableNames.FirstOrDefault(
            name => string.Equals(name, selectedTable, StringComparison.OrdinalIgnoreCase));
        if (match == null) {
            throw new KeyNotFoundException($"Table '{selectedTable}' was not found.");
        }

        return new[] { match };
    }

    private static ExcelReadOptions CreateExcelOptions(TabularReadOptions options) =>
        new() {
            MaxInputBytes = options.MaxInputBytes,
            Culture = options.Culture,
            NumericAsDecimal = options.NumericAsDecimal,
            TreatDatesUsingNumberFormat = options.TreatDatesUsingNumberFormat,
            UseCachedFormulaResult = options.UseCachedFormulaResult
        };

    private static TabularReader OpenExcelBinaryPath(string path, TabularReadOptions options) {
        ExcelReadOptions excelOptions = CreateExcelOptions(options);
        XlsbTabularWorkbook owner = XlsbTabularWorkbook.Open(
            path,
            excelOptions,
            options.CancellationToken);
        return CreateExcelBinaryReader(owner, options, excelOptions);
    }

    private static TabularReader OpenExcelBinaryStream(Stream stream, TabularReadOptions options) {
        ExcelReadOptions excelOptions = CreateExcelOptions(options);
        XlsbTabularWorkbook owner = XlsbTabularWorkbook.Open(
            stream,
            excelOptions,
            options.CancellationToken);
        return CreateExcelBinaryReader(owner, options, excelOptions);
    }

    private static TabularReader CreateExcelBinaryReader(
        XlsbTabularWorkbook owner,
        TabularReadOptions options,
        ExcelReadOptions excelOptions) {
        IReadOnlyList<string> selectedTables = SelectTables(owner.TableNames, options.TableName);
        try {
            return new TabularReader(
                TabularFormat.ExcelBinary,
                selectedTables,
                tableIndex => owner.OpenTable(
                    selectedTables[tableIndex],
                    options.HasHeaderRow,
                    excelOptions,
                    options.CancellationToken),
                owner,
                options.Culture);
        } catch {
            owner.Dispose();
            throw;
        }
    }
}
