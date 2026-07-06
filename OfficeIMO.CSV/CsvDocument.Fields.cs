#nullable enable

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
#if NET8_0_OR_GREATER
    /// <summary>
    /// Reads CSV data rows from a file in a single pass, applying header discovery while avoiding string materialization for unquoted data fields.
    /// </summary>
    /// <typeparam name="TVisitor">Struct visitor type receiving each data row and field.</typeparam>
    /// <param name="path">Source CSV path.</param>
    /// <param name="rowVisitor">Visitor receiving normalized headers and transient data field spans.</param>
    /// <param name="options">Optional load settings.</param>
    public static void ReadRowFieldSpans<TVisitor>(string path, ref TVisitor rowVisitor, CsvLoadOptions? options = null)
        where TVisitor : struct, ICsvRowFieldSpanVisitor
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options ??= new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var readerFactory = () => new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: FileBufferSize);
        var resolvedOptions = ResolveLoadOptions(readerFactory, options);
        using var reader = readerFactory();
        ReadRowFieldSpans(reader, ref rowVisitor, resolvedOptions);
    }

    /// <summary>
    /// Reads CSV data rows from a reader in a single pass, applying header discovery while avoiding string materialization for unquoted data fields.
    /// </summary>
    /// <typeparam name="TVisitor">Struct visitor type receiving each data row and field.</typeparam>
    /// <param name="reader">Source text reader.</param>
    /// <param name="rowVisitor">Visitor receiving normalized headers and transient data field spans.</param>
    /// <param name="options">Optional load settings.</param>
    public static void ReadRowFieldSpans<TVisitor>(TextReader reader, ref TVisitor rowVisitor, CsvLoadOptions? options = null)
        where TVisitor : struct, ICsvRowFieldSpanVisitor
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        options ??= new CsvLoadOptions();
        if (options.DetectDelimiter)
        {
            var text = reader.ReadToEnd();
            var resolvedOptions = ResolveLoadOptions(() => new StringReader(text), options);
            using var bufferedReader = new StringReader(text);
            ReadRowFieldSpans(bufferedReader, ref rowVisitor, resolvedOptions);
            return;
        }

        var explicitHeader = NormalizeExplicitHeader(options);
        if (explicitHeader is not null)
        {
            ReadRowFieldSpansWithHeader(reader, explicitHeader, options, GetInitialRecordsToSkip(options), ref rowVisitor);
            return;
        }

        var visitor = new CsvHeaderAwareFieldSpanVisitor<TVisitor>(rowVisitor, options, firstRecordIsData: !options.HasHeaderRow);
        CsvParser.ReadFieldSpans(reader, options, GetInitialRecordsToSkip(options), ref visitor);
        visitor.Complete();
        rowVisitor = visitor.RowVisitor;
    }

    /// <summary>
    /// Reads CSV fields from a file in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="fieldAction">Action receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpans(string path, CsvFieldSpanAction fieldAction, CsvLoadOptions? options = null)
    {
        if (fieldAction == null)
        {
            throw new ArgumentNullException(nameof(fieldAction));
        }

        var visitor = new CsvFieldSpanActionVisitor(fieldAction);
        ReadFieldSpans(path, ref visitor, options);
    }

    /// <summary>
    /// Reads CSV fields from a file in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <typeparam name="TVisitor">Struct visitor type receiving each field.</typeparam>
    /// <param name="path">Source CSV path.</param>
    /// <param name="fieldVisitor">Visitor receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpans<TVisitor>(string path, ref TVisitor fieldVisitor, CsvLoadOptions? options = null)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options = CreateRawRecordOptions(options);
        var readerFactory = () => CsvFile.OpenTextReader(path, options, FileBufferSize);
        var resolvedOptions = ResolveLoadOptions(readerFactory, options, useHeaderDiscoveryForDelimiterDetection: false);
        using var reader = readerFactory();
        ReadFieldSpans(reader, ref fieldVisitor, resolvedOptions);
    }

    /// <summary>
    /// Reads CSV fields from a reader in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <param name="reader">Source text reader.</param>
    /// <param name="fieldAction">Action receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpans(TextReader reader, CsvFieldSpanAction fieldAction, CsvLoadOptions? options = null)
    {
        if (fieldAction == null)
        {
            throw new ArgumentNullException(nameof(fieldAction));
        }

        var visitor = new CsvFieldSpanActionVisitor(fieldAction);
        ReadFieldSpans(reader, ref visitor, options);
    }

    /// <summary>
    /// Reads CSV fields from a reader in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <typeparam name="TVisitor">Struct visitor type receiving each field.</typeparam>
    /// <param name="reader">Source text reader.</param>
    /// <param name="fieldVisitor">Visitor receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpans<TVisitor>(TextReader reader, ref TVisitor fieldVisitor, CsvLoadOptions? options = null)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        options = CreateRawRecordOptions(options);
        if (options.DetectDelimiter)
        {
            var text = reader.ReadToEnd();
            var resolvedOptions = ResolveLoadOptions(() => new StringReader(text), options, useHeaderDiscoveryForDelimiterDetection: false);
            using var bufferedReader = new StringReader(text);
            ReadFieldSpans(bufferedReader, ref fieldVisitor, resolvedOptions);
            return;
        }

        CsvParser.ReadFieldSpans(reader, options, GetInitialRecordsToSkip(options), ref fieldVisitor);
    }

    /// <summary>
    /// Reads CSV fields from text in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <param name="text">Source CSV text.</param>
    /// <param name="fieldAction">Action receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpansFromText(string text, CsvFieldSpanAction fieldAction, CsvLoadOptions? options = null)
    {
        if (fieldAction == null)
        {
            throw new ArgumentNullException(nameof(fieldAction));
        }

        var visitor = new CsvFieldSpanActionVisitor(fieldAction);
        ReadFieldSpansFromText(text, ref visitor, options);
    }

    /// <summary>
    /// Reads CSV fields from text in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <typeparam name="TVisitor">Struct visitor type receiving each field.</typeparam>
    /// <param name="text">Source CSV text.</param>
    /// <param name="fieldVisitor">Visitor receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpansFromText<TVisitor>(string text, ref TVisitor fieldVisitor, CsvLoadOptions? options = null)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (text == null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        ReadFieldSpans(text.AsSpan(), ref fieldVisitor, options);
    }

    /// <summary>
    /// Reads CSV fields from text in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <typeparam name="TVisitor">Struct visitor type receiving each field.</typeparam>
    /// <param name="text">Source CSV text.</param>
    /// <param name="fieldVisitor">Visitor receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpans<TVisitor>(ReadOnlySpan<char> text, ref TVisitor fieldVisitor, CsvLoadOptions? options = null)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        options = CreateRawRecordOptions(options);
        if (options.DetectDelimiter)
        {
            var sourceText = text.ToString();
            var resolvedOptions = ResolveLoadOptions(() => new StringReader(sourceText), options, useHeaderDiscoveryForDelimiterDetection: false);
            CsvParser.ReadFieldSpans(sourceText.AsSpan(), resolvedOptions, GetInitialRecordsToSkip(resolvedOptions), ref fieldVisitor);
            return;
        }

        CsvParser.ReadFieldSpans(text, options, GetInitialRecordsToSkip(options), ref fieldVisitor);
    }

    private static void ReadRowFieldSpansWithHeader<TVisitor>(
        TextReader reader,
        IReadOnlyList<string> header,
        CsvLoadOptions options,
        int recordsToSkip,
        ref TVisitor rowVisitor)
        where TVisitor : struct, ICsvRowFieldSpanVisitor
    {
        var visitor = new CsvHeaderAwareFieldSpanVisitor<TVisitor>(rowVisitor, options, firstRecordIsData: false, header: header);
        CsvParser.ReadFieldSpans(reader, options, recordsToSkip, ref visitor);
        visitor.Complete();
        rowVisitor = visitor.RowVisitor;
    }

    private struct CsvHeaderAwareFieldSpanVisitor<TVisitor> : ICsvFieldSpanVisitor
        where TVisitor : struct, ICsvRowFieldSpanVisitor
    {
        private readonly CsvLoadOptions _options;
        private readonly List<string> _headerFields;
        private readonly bool _firstRecordIsData;
        private TVisitor _rowVisitor;
        private IReadOnlyList<string>? _header;
        private int _currentRecordIndex;
        private int _currentFieldCount;
        private int _rowIndex;
        private bool _hasCurrentRecord;
        private bool _hasRowStarted;
        private bool _needsHeader;

        public CsvHeaderAwareFieldSpanVisitor(TVisitor rowVisitor, CsvLoadOptions options, bool firstRecordIsData, IReadOnlyList<string>? header = null)
        {
            _rowVisitor = rowVisitor;
            _options = options;
            _firstRecordIsData = firstRecordIsData;
            _header = header;
            _headerFields = new List<string>(64);
            _currentRecordIndex = 0;
            _currentFieldCount = 0;
            _rowIndex = 0;
            _hasCurrentRecord = false;
            _hasRowStarted = false;
            _needsHeader = header is null;
        }

        public readonly TVisitor RowVisitor => _rowVisitor;

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            BeginOrAdvanceRecord(recordIndex);
            _currentFieldCount++;
            if (_needsHeader)
            {
                _headerFields.Add(value.ToString());
                return;
            }

            var header = _header!;
            if (!_hasRowStarted)
            {
                _rowVisitor.BeginRow(header, _rowIndex);
                _hasRowStarted = true;
            }

            if (fieldIndex < header.Count)
            {
                _rowVisitor.VisitField(_rowIndex, fieldIndex, value);
            }
        }

        public void VisitFieldValue(int recordIndex, int fieldIndex, string value)
        {
            BeginOrAdvanceRecord(recordIndex);
            _currentFieldCount++;
            if (_needsHeader)
            {
                _headerFields.Add(value);
                return;
            }

            var header = _header!;
            if (!_hasRowStarted)
            {
                _rowVisitor.BeginRow(header, _rowIndex);
                _hasRowStarted = true;
            }

            if (fieldIndex < header.Count)
            {
                _rowVisitor.VisitFieldValue(_rowIndex, fieldIndex, value);
            }
        }

        private void BeginOrAdvanceRecord(int recordIndex)
        {
            if (!_hasCurrentRecord)
            {
                BeginRecord(recordIndex);
            }
            else if (recordIndex != _currentRecordIndex)
            {
                CompleteCurrentRecord();
                BeginRecord(recordIndex);
            }
        }

        public void Complete()
        {
            if (_hasCurrentRecord)
            {
                CompleteCurrentRecord();
            }
        }

        private void BeginRecord(int recordIndex)
        {
            _hasCurrentRecord = true;
            _hasRowStarted = false;
            _currentRecordIndex = recordIndex;
            _currentFieldCount = 0;
            if (_needsHeader)
            {
                _headerFields.Clear();
            }
        }

        private void CompleteCurrentRecord()
        {
            if (_needsHeader)
            {
                if (_firstRecordIsData)
                {
                    _header = GenerateDefaultHeader(_headerFields.Count);
                    _needsHeader = false;
                    EmitBufferedFirstDataRow();
                    return;
                }

                _header = ResolveHeader(_headerFields, _options);
                _needsHeader = false;
                return;
            }

            var header = _header!;
            if (_options.ColumnCountMismatchPolicy == CsvColumnCountMismatchPolicy.Strict &&
                _currentFieldCount != header.Count)
            {
                throw new CsvException($"Row contains {_currentFieldCount} values but header defines {header.Count} columns.");
            }

            if (!_hasRowStarted)
            {
                _rowVisitor.BeginRow(header, _rowIndex);
            }

            _rowVisitor.EndRow(_rowIndex, _currentFieldCount);
            _rowIndex++;
        }

        private void EmitBufferedFirstDataRow()
        {
            var header = _header!;
            _rowVisitor.BeginRow(header, _rowIndex);
            for (var i = 0; i < _headerFields.Count; i++)
            {
                _rowVisitor.VisitFieldValue(_rowIndex, i, _headerFields[i]);
            }

            _rowVisitor.EndRow(_rowIndex, _headerFields.Count);
            _rowIndex++;
        }

        private static IReadOnlyList<string> ResolveHeader(IReadOnlyList<string> fields, CsvLoadOptions options)
        {
            if (TryGetW3CFieldsHeader(fields, options, out var w3cHeader))
            {
                return w3cHeader;
            }

            return NormalizeParsedHeader(fields, options);
        }
    }
#endif
}
