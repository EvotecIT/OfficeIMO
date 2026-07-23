#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    private readonly struct CsvLine
    {
        public CsvLine(string text, string separator)
        {
            Text = text;
            Separator = separator;
        }

        public string Text { get; }

        public string Separator { get; }
    }

    internal readonly struct CsvParsedRecord
    {
        public CsvParsedRecord(IReadOnlyList<string> values, bool startsWithCommentCharacter)
        {
            Values = values;
            StartsWithCommentCharacter = startsWithCommentCharacter;
        }

        public IReadOnlyList<string> Values { get; }

        public bool StartsWithCommentCharacter { get; }
    }

    public static IEnumerable<string[]> Parse(TextReader reader, CsvLoadOptions options)
    {
        if (UsesTextDelimiter(options))
        {
            return ParseLineOrQuotedTextDelimiter(reader, options);
        }

        return ParseLineOrQuoted(reader, options);
    }

    public static void ReadRecords(TextReader reader, CsvLoadOptions options, Action<string[]> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        if (UsesTextDelimiter(options))
        {
            ReadLineOrQuotedTextDelimiter(reader, options, recordAction);
            return;
        }

        ReadLineOrQuoted(reader, options, recordAction);
    }

    internal static IEnumerable<CsvParsedRecord> ParseWithMetadata(TextReader reader, CsvLoadOptions options)
    {
        if (UsesTextDelimiter(options))
        {
            return ParseLineOrQuotedTextDelimiterWithMetadata(reader, options);
        }

        return ParseLineOrQuotedWithMetadata(reader, options);
    }

    internal static void ReadRecordsWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        var records = UsesTextDelimiter(options)
            ? ParseLineOrQuotedTextDelimiterWithMetadata(reader, options)
            : ParseLineOrQuotedWithMetadata(reader, options);

        foreach (var record in records)
        {
            recordAction(record);
        }
    }

    public static void ReadRecordsReusable(TextReader reader, CsvLoadOptions options, Action<IReadOnlyList<string>> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        if (UsesTextDelimiter(options))
        {
            ReadLineOrQuotedTextDelimiterReusable(reader, options, recordAction);
            return;
        }

        ReadLineOrQuotedReusable(reader, options, recordAction);
    }

    internal static void ReadRecordsReusableWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        if (UsesTextDelimiter(options))
        {
            ReadLineOrQuotedTextDelimiterReusableWithMetadata(reader, options, recordAction);
            return;
        }

        ReadLineOrQuotedReusableWithMetadata(reader, options, recordAction);
    }

#if NET8_0_OR_GREATER
    internal static void ReadFieldSpans(TextReader reader, CsvLoadOptions options, int recordsToSkip, CsvFieldSpanAction fieldAction)
    {
        if (fieldAction == null)
        {
            throw new ArgumentNullException(nameof(fieldAction));
        }

        var visitor = new CsvFieldSpanActionVisitor(fieldAction);
        ReadFieldSpans(reader, options, recordsToSkip, ref visitor);
    }

    internal static void ReadFieldSpans<TVisitor>(TextReader reader, CsvLoadOptions options, int recordsToSkip, ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (recordsToSkip < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(recordsToSkip), "SkipInitialRecords cannot be negative.");
        }

        if (HasFieldLengthLimits(options))
        {
            ReadFieldSpansMaterialized(reader, options, recordsToSkip, ref fieldVisitor);
            return;
        }

        if (UsesTextDelimiter(options))
        {
            ReadFieldSpansTextDelimiter(reader, options, recordsToSkip, ref fieldVisitor);
            return;
        }

        ReadFieldSpansLineOrQuoted(reader, options, recordsToSkip, ref fieldVisitor);
    }

    private static bool HasFieldLengthLimits(CsvLoadOptions options) =>
        options.MaxFieldLength.HasValue || options.MaxQuotedFieldLength.HasValue;

    private static bool NeedsLogicalCommentSkipping(CsvLoadOptions options) =>
        options.SkipCommentRows ||
        (options.HasHeaderRow &&
            options.Header is null &&
            options.SkipCommentRowsBeforeHeader);

    private static void ReadFieldSpansMaterialized<TVisitor>(TextReader reader, CsvLoadOptions options, int recordsToSkip, ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var recordIndex = 0;
        var projectedFieldVisitor = fieldVisitor as ICsvProjectedFieldSpanVisitor;
        foreach (var record in ParseReusable(reader, options))
        {
            if (recordsToSkip > 0)
            {
                recordsToSkip--;
                continue;
            }

            VisitParsedFields(record, recordIndex, projectedFieldVisitor, ref fieldVisitor);
            recordIndex++;
        }
    }
#endif

    private static void ReadLineOrQuoted(TextReader reader, CsvLoadOptions options, Action<string[]> recordAction)
    {
        var delimiter = GetDelimiterChar(options);
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();
        var stringCache = CreateStringCache(options);
        using var lineReader = new CsvLineReader(reader);

        while (ReadLineWithSeparator(lineReader, pendingLines, out var lineSeparator) is { } line)
        {
            ThrowIfCancellationRequested(options);
            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (TrySkipCommentRecordBeforeParsing(lineReader, pendingLines, startsWithCommentCharacter, line, lineSeparator, options, emittedRecordCount, ref lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, out var record))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount) &&
                    ShouldEmitRecord(record, allowEmpty))
                {
                    if (!TryPrepareParsedRecord(record, options, lineNumber, quotedRecord: false, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    recordAction(record);
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            try
            {
                if (!TryParseQuotedRecordWithContinuations(lineReader, pendingLines, line, lineSeparator, delimiter, trim, strictQuotes, options, ref lineNumber, out fields))
                {
                    throw new CsvParseException("Unterminated quoted field.", lineNumber);
                }
            }
            catch (CsvParseException ex) when (HandleParseError(options, ex, lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    if (!TryPrepareParsedRecord(fields, options, lineNumber, quotedRecord: true, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    recordAction(fields);
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }
            }

            lineNumber++;
        }
    }

    private static IEnumerable<string[]> ParseLineOrQuoted(TextReader reader, CsvLoadOptions options)
    {
        foreach (var record in ParseLineOrQuotedWithMetadata(reader, options))
        {
            if (record.Values is string[] fields)
            {
                yield return fields;
            }
        }
    }

    private static IEnumerable<CsvParsedRecord> ParseLineOrQuotedWithMetadata(TextReader reader, CsvLoadOptions options)
    {
        var delimiter = GetDelimiterChar(options);
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();
        var stringCache = CreateStringCache(options);
        using var lineReader = new CsvLineReader(reader);

        while (ReadLineWithSeparator(lineReader, pendingLines, out var lineSeparator) is { } line)
        {
            ThrowIfCancellationRequested(options);
            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (TrySkipCommentRecordBeforeParsing(lineReader, pendingLines, startsWithCommentCharacter, line, lineSeparator, options, emittedRecordCount, ref lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, out var record))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount) &&
                    ShouldEmitRecord(record, allowEmpty))
                {
                    if (!TryPrepareParsedRecord(record, options, lineNumber, quotedRecord: false, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    yield return new CsvParsedRecord(record, startsWithCommentCharacter);
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            try
            {
                if (!TryParseQuotedRecordWithContinuations(lineReader, pendingLines, line, lineSeparator, delimiter, trim, strictQuotes, options, ref lineNumber, out fields))
                {
                    throw new CsvParseException("Unterminated quoted field.", lineNumber);
                }
            }
            catch (CsvParseException ex) when (HandleParseError(options, ex, lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    if (!TryPrepareParsedRecord(fields, options, lineNumber, quotedRecord: true, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    yield return new CsvParsedRecord(fields, startsWithCommentCharacter);
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }
            }

            lineNumber++;
        }
    }

    private static void ReadLineOrQuotedReusable(TextReader reader, CsvLoadOptions options, Action<IReadOnlyList<string>> recordAction)
    {
        var delimiter = GetDelimiterChar(options);
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var reusableRecord = new List<string>(64);
        var reusableQuotedRecord = new List<string>(64);
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();
        var stringCache = CreateStringCache(options);
        using var lineReader = new CsvLineReader(reader);

        while (pendingLines.Count > 0 || true)
        {
            ThrowIfCancellationRequested(options);
            string? fastLine = null;
            string lineSeparator;
            CsvLineReadResult readResult;
            if (pendingLines.Count == 0)
            {
                readResult = lineReader.ReadUnquotedRecordOrLine(delimiter, trim, options.CommentCharacter, reusableRecord, out fastLine, out lineSeparator);
            }
            else
            {
                lineSeparator = string.Empty;
                readResult = CsvLineReadResult.Line;
            }

            if (readResult == CsvLineReadResult.EndOfReader)
            {
                break;
            }

            if (readResult == CsvLineReadResult.UnquotedRecord)
            {
                if (ShouldEmitRecord(reusableRecord, allowEmpty))
                {
                    if (!TryPrepareParsedRecord(reusableRecord, options, lineNumber, quotedRecord: false, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    recordAction(reusableRecord);
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }

                lineNumber++;
                continue;
            }

            var line = pendingLines.Count > 0
                ? ReadLineWithSeparator(lineReader, pendingLines, out lineSeparator)
                : fastLine;
            if (line is null)
            {
                break;
            }

            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (TrySkipCommentRecordBeforeParsing(lineReader, pendingLines, startsWithCommentCharacter, line, lineSeparator, options, emittedRecordCount, ref lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, reusableRecord))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount) &&
                    ShouldEmitRecord(reusableRecord, allowEmpty))
                {
                    if (!TryPrepareParsedRecord(reusableRecord, options, lineNumber, quotedRecord: false, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    recordAction(reusableRecord);
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }

                lineNumber++;
                continue;
            }

            try
            {
                if (!TryParseQuotedRecordWithContinuations(lineReader, pendingLines, line, lineSeparator, delimiter, trim, strictQuotes, options, ref lineNumber, reusableQuotedRecord))
                {
                    throw new CsvParseException("Unterminated quoted field.", lineNumber);
                }
            }
            catch (CsvParseException ex) when (HandleParseError(options, ex, lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (ShouldEmitRecord(reusableQuotedRecord, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    if (!TryPrepareParsedRecord(reusableQuotedRecord, options, lineNumber, quotedRecord: true, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    recordAction(reusableQuotedRecord);
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }
            }

            lineNumber++;
        }
    }

    private static void ReadLineOrQuotedReusableWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        var delimiter = GetDelimiterChar(options);
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var reusableRecord = new List<string>(64);
        var reusableQuotedRecord = new List<string>(64);
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();
        var stringCache = CreateStringCache(options);
        using var lineReader = new CsvLineReader(reader);

        while (pendingLines.Count > 0 || true)
        {
            ThrowIfCancellationRequested(options);
            string? fastLine = null;
            string lineSeparator;
            CsvLineReadResult readResult;
            if (pendingLines.Count == 0)
            {
                readResult = lineReader.ReadUnquotedRecordOrLine(delimiter, trim, options.CommentCharacter, reusableRecord, out fastLine, out lineSeparator);
            }
            else
            {
                lineSeparator = string.Empty;
                readResult = CsvLineReadResult.Line;
            }

            if (readResult == CsvLineReadResult.EndOfReader)
            {
                break;
            }

            if (readResult == CsvLineReadResult.UnquotedRecord)
            {
                if (ShouldEmitRecord(reusableRecord, allowEmpty))
                {
                    if (!TryPrepareParsedRecord(reusableRecord, options, lineNumber, quotedRecord: false, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    recordAction(new CsvParsedRecord(reusableRecord, startsWithCommentCharacter: false));
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }

                lineNumber++;
                continue;
            }

            var line = pendingLines.Count > 0
                ? ReadLineWithSeparator(lineReader, pendingLines, out lineSeparator)
                : fastLine;
            if (line is null)
            {
                break;
            }

            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (TrySkipCommentRecordBeforeParsing(lineReader, pendingLines, startsWithCommentCharacter, line, lineSeparator, options, emittedRecordCount, ref lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (TrySplitUnquotedRecord(line, delimiter, trim, reusableRecord))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount) &&
                    ShouldEmitRecord(reusableRecord, allowEmpty))
                {
                    if (!TryPrepareParsedRecord(reusableRecord, options, lineNumber, quotedRecord: false, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    recordAction(new CsvParsedRecord(reusableRecord, startsWithCommentCharacter));
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }

                lineNumber++;
                continue;
            }

            try
            {
                if (!TryParseQuotedRecordWithContinuations(lineReader, pendingLines, line, lineSeparator, delimiter, trim, strictQuotes, options, ref lineNumber, reusableQuotedRecord))
                {
                    throw new CsvParseException("Unterminated quoted field.", lineNumber);
                }
            }
            catch (CsvParseException ex) when (HandleParseError(options, ex, lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (ShouldEmitRecord(reusableQuotedRecord, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    if (!TryPrepareParsedRecord(reusableQuotedRecord, options, lineNumber, quotedRecord: true, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    recordAction(new CsvParsedRecord(reusableQuotedRecord, startsWithCommentCharacter));
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }
            }

            lineNumber++;
        }
    }

#if NET8_0_OR_GREATER
    private static void ReadFieldSpansLineOrQuoted<TVisitor>(TextReader reader, CsvLoadOptions options, int recordsToSkip, ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var delimiter = GetDelimiterChar(options);
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var emittedRecordCount = 0;
        var recordIndex = 0;
        var reusableQuotedRecord = new List<string>(64);
        var pendingLines = new Queue<CsvLine>();
        var projectedFieldVisitor = fieldVisitor as ICsvProjectedFieldSpanVisitor;
        using var lineReader = new CsvLineReader(reader);

        while (pendingLines.Count > 0 || true)
        {
            ThrowIfCancellationRequested(options);
            string? fastLine = null;
            string lineSeparator;
            CsvLineReadResult readResult;
            if (pendingLines.Count == 0)
            {
                readResult = lineReader.ReadUnquotedFieldSpansOrLine(
                    delimiter,
                    trim,
                    options.CommentCharacter,
                    allowEmpty,
                    recordsToSkip == 0,
                    recordIndex,
                    projectedFieldVisitor,
                    ref fieldVisitor,
                    out var fieldCount,
                    out var isEmptyRecord,
                    out fastLine,
                    out lineSeparator);

                if (readResult == CsvLineReadResult.EndOfReader)
                {
                    break;
                }

                if (readResult == CsvLineReadResult.UnquotedRecord)
                {
                    if (fieldCount != 0 && (allowEmpty || !isEmptyRecord))
                    {
                        if (recordsToSkip > 0)
                        {
                            recordsToSkip--;
                        }
                        else
                        {
                            recordIndex++;
                        }

                        emittedRecordCount++;
                        ReportProgress(options, emittedRecordCount, lineNumber);
                    }

                    lineNumber++;
                    continue;
                }
            }
            else
            {
                lineSeparator = string.Empty;
                readResult = CsvLineReadResult.Line;
            }

            var line = pendingLines.Count > 0
                ? ReadLineWithSeparator(lineReader, pendingLines, out lineSeparator)
                : fastLine;
            if (line is null)
            {
                break;
            }

            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (TrySkipCommentRecordBeforeParsing(lineReader, pendingLines, startsWithCommentCharacter, line, lineSeparator, options, emittedRecordCount, ref lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (line.IndexOf('"') < 0 && TrySplitUnquotedRecord(line, delimiter, trim, out var record))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount) &&
                    ShouldEmitRecord(record, allowEmpty))
                {
                    if (recordsToSkip > 0)
                    {
                        recordsToSkip--;
                    }
                    else
                    {
                        VisitParsedFields(record, recordIndex, projectedFieldVisitor, ref fieldVisitor);
                        recordIndex++;
                    }

                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }

                lineNumber++;
                continue;
            }

            try
            {
                if (!TryParseQuotedRecordContinuations(
                        lineReader,
                        pendingLines,
                        line,
                        lineSeparator,
                        delimiter,
                        trim,
                        strictQuotes,
                        reusableQuotedRecord,
                        ref lineNumber) &&
                    !TryParseQuotedRecord(line, delimiter, trim, strictQuotes, lineNumber, reusableQuotedRecord))
                {
                    throw new CsvParseException("Unterminated quoted field.", lineNumber);
                }
            }
            catch (CsvParseException ex) when (HandleParseError(options, ex, lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (ShouldEmitRecord(reusableQuotedRecord, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    if (recordsToSkip > 0)
                    {
                        recordsToSkip--;
                    }
                    else
                    {
                        VisitParsedFields(reusableQuotedRecord, recordIndex, projectedFieldVisitor, ref fieldVisitor);
                        recordIndex++;
                    }

                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }
            }

            lineNumber++;
        }
    }

    private static void VisitParsedFields<TVisitor>(
        IReadOnlyList<string> fields,
        int recordIndex,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        for (var fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
        {
            if (CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
            {
                fieldVisitor.VisitFieldValue(recordIndex, fieldIndex, fields[fieldIndex]);
            }
        }
    }
#endif

    private static bool TrySplitUnquotedRecord(string line, char delimiter, bool trim, out string[] fields)
    {
        var fieldCount = 1;
        for (var i = 0; i < line.Length; i++)
        {
            var value = line[i];
            if (value == '"')
            {
                fields = Array.Empty<string>();
                return false;
            }

            if (value == delimiter)
            {
                fieldCount++;
            }
        }

        fields = new string[fieldCount];
        var fieldIndex = 0;
        var start = 0;
        for (var i = 0; i < line.Length; i++)
        {
            if (line[i] != delimiter)
            {
                continue;
            }

            fields[fieldIndex] = GetUnquotedField(line, start, i - start, trim);
            fieldIndex++;
            start = i + 1;
        }

        fields[fieldIndex] = GetUnquotedField(line, start, line.Length - start, trim);
        return true;
    }

    private static bool TrySplitUnquotedRecord(string line, char delimiter, bool trim, List<string> fields)
    {
        fields.Clear();
        var start = 0;
        for (var i = 0; i < line.Length; i++)
        {
            var value = line[i];
            if (value == '"')
            {
                fields.Clear();
                return false;
            }

            if (value != delimiter)
            {
                continue;
            }

            fields.Add(GetUnquotedField(line, start, i - start, trim));
            start = i + 1;
        }

        fields.Add(GetUnquotedField(line, start, line.Length - start, trim));
        return true;
    }

    private static string GetUnquotedField(string line, int start, int length, bool trim)
    {
        if (length == 0)
        {
            return string.Empty;
        }

        if (!trim)
        {
            return line.Substring(start, length);
        }

        var end = start + length - 1;
        while (start <= end && char.IsWhiteSpace(line[start]))
        {
            start++;
        }

        while (end >= start && char.IsWhiteSpace(line[end]))
        {
            end--;
        }

        if (end < start)
        {
            return string.Empty;
        }

        return line.Substring(start, end - start + 1);
    }

    private static bool ShouldSkipCommentRecord(bool startsWithCommentCharacter, string firstLine, CsvLoadOptions options, int emittedRecordCount)
    {
        if (!options.SkipCommentRows || !startsWithCommentCharacter)
        {
            return false;
        }

        return !CanReadW3CFieldsHeader(options, emittedRecordCount) || !IsW3CFieldsLine(firstLine, options);
    }

    private static bool TrySkipCommentRecordBeforeParsing(CsvLineReader reader, Queue<CsvLine> pendingLines, bool startsWithCommentCharacter, string firstLine, string firstLineSeparator, CsvLoadOptions options, int emittedRecordCount, ref int lineNumber)
    {
        if (!ShouldSkipCommentRecordBeforeParsing(startsWithCommentCharacter, firstLine, options, emittedRecordCount))
        {
            return false;
        }

        var delimiter = GetDelimiterChar(options);
        if (TryParseQuotedRecordLenient(firstLine, delimiter, options.TrimWhitespace, out _))
        {
            return true;
        }

        var continuations = new List<CsvLine>();
        var state = new QuotedRecordState();
        UpdateQuotedRecordState(firstLine, delimiter, options.TrimWhitespace, ref state);
        while (true)
        {
            var next = ReadLineWithSeparator(reader, pendingLines, out var nextSeparator);
            if (next == null)
            {
                EnqueueContinuations(pendingLines, continuations);
                return true;
            }

            continuations.Add(new CsvLine(next, nextSeparator));
            UpdateQuotedRecordState(next, delimiter, options.TrimWhitespace, ref state);
            if (!state.InQuotes)
            {
                lineNumber += continuations.Count;
                return true;
            }

            if (!LooksLikeDelimitedRawComment(firstLine, delimiter) && LooksLikeDelimitedRawComment(next, delimiter))
            {
                EnqueueContinuations(pendingLines, continuations);
                return true;
            }
        }
    }

    private static void EnqueueContinuations(Queue<CsvLine> pendingLines, List<CsvLine> continuations)
    {
        foreach (var continuation in continuations)
        {
            pendingLines.Enqueue(continuation);
        }
    }

    private static bool LooksLikeDelimitedRawComment(string line, char delimiter) =>
        line.IndexOf(delimiter) >= 0 ||
        line.IndexOf(',') >= 0 ||
        line.IndexOf(';') >= 0 ||
        line.IndexOf('|') >= 0 ||
        line.IndexOf('\t') >= 0;

    private static bool ShouldSkipCommentRecordBeforeParsing(bool startsWithCommentCharacter, string firstLine, CsvLoadOptions options, int emittedRecordCount)
    {
        if (!startsWithCommentCharacter)
        {
            return false;
        }

        var canReadW3CFieldsHeader = CanReadW3CFieldsHeader(options, emittedRecordCount) && IsW3CFieldsLine(firstLine, options);
        if (canReadW3CFieldsHeader)
        {
            return false;
        }

        return options.SkipCommentRows ||
            (options.HasHeaderRow &&
                options.Header is null &&
                options.SkipCommentRowsBeforeHeader &&
                emittedRecordCount <= GetParserInitialRecordsToSkip(options));
    }

    private static bool IsRawCommentLine(string line, CsvLoadOptions options) =>
        line.Length > 0 && line[0] == options.CommentCharacter;

    private static bool CanReadW3CFieldsHeader(CsvLoadOptions options, int emittedRecordCount) =>
        emittedRecordCount <= GetParserInitialRecordsToSkip(options) &&
        options.HasHeaderRow &&
        options.Header is null &&
        options.RecognizeW3CFieldsHeader;

    private static int GetParserInitialRecordsToSkip(CsvLoadOptions options)
    {
        if (options.SkipInitialRecords < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "SkipInitialRecords cannot be negative.");
        }

        return options.SkipInitialRecords;
    }

    private static bool IsW3CFieldsLine(string line, CsvLoadOptions options) =>
        options.RecognizeW3CFieldsHeader && line.StartsWith("#Fields:", StringComparison.OrdinalIgnoreCase);

    private static string? ReadLineWithSeparator(CsvLineReader reader, Queue<CsvLine> pendingLines, out string separator)
    {
        if (pendingLines.Count > 0)
        {
            var pending = pendingLines.Dequeue();
            separator = pending.Separator;
            return pending.Text;
        }

        return reader.ReadLine(out separator);
    }

    private static bool ShouldEmitRecord(IReadOnlyList<string> fields, bool allowEmpty)
    {
        return fields.Count != 0 &&
            (allowEmpty || fields.Count != 1 || fields[0].Length != 0);
    }

    private static Dictionary<string, string>? CreateStringCache(CsvLoadOptions options) =>
        options.InternStrings ? new Dictionary<string, string>(StringComparer.Ordinal) : null;

    private static void ThrowIfCancellationRequested(CsvLoadOptions options)
    {
        if (options.CancellationToken.CanBeCanceled)
        {
            options.CancellationToken.ThrowIfCancellationRequested();
        }
    }

    private static void ReportProgress(CsvLoadOptions options, long emittedRecordCount, int lineNumber)
    {
        if (options.ProgressCallback is null || options.ProgressReportInterval <= 0)
        {
            return;
        }

        if (emittedRecordCount % options.ProgressReportInterval == 0)
        {
            options.ProgressCallback(new CsvProgress(emittedRecordCount, lineNumber));
        }
    }

    private static bool HandleParseError(CsvLoadOptions options, CsvParseException exception, int lineNumber)
    {
        if (options.CollectParseErrors)
        {
            var errors = options.ParseErrors ??= new List<CsvParseError>();
            errors.Add(new CsvParseError(exception.LineNumber ?? lineNumber, exception.Message, exception));
            if (options.MaxParseErrors >= 0 && errors.Count > options.MaxParseErrors)
            {
                throw new CsvParseException($"CSV parse error limit of {options.MaxParseErrors} was exceeded.", lineNumber, exception);
            }
        }

        if (options.ParseErrorAction == CsvParseErrorAction.SkipRow)
        {
            return true;
        }

        throw exception;
    }

    private static void PrepareParsedRecord(string[] fields, CsvLoadOptions options, int lineNumber, bool quotedRecord, Dictionary<string, string>? stringCache)
    {
        for (var i = 0; i < fields.Length; i++)
        {
            fields[i] = PrepareParsedField(fields[i], options, lineNumber, quotedRecord, stringCache);
        }
    }

    private static void PrepareParsedRecord(List<string> fields, CsvLoadOptions options, int lineNumber, bool quotedRecord, Dictionary<string, string>? stringCache)
    {
        for (var i = 0; i < fields.Count; i++)
        {
            fields[i] = PrepareParsedField(fields[i], options, lineNumber, quotedRecord, stringCache);
        }
    }

    private static bool TryPrepareParsedRecord(string[] fields, CsvLoadOptions options, int lineNumber, bool quotedRecord, Dictionary<string, string>? stringCache)
    {
        try
        {
            PrepareParsedRecord(fields, options, lineNumber, quotedRecord, stringCache);
            return true;
        }
        catch (CsvParseException ex) when (HandleParseError(options, ex, lineNumber))
        {
            return false;
        }
    }

    private static bool TryPrepareParsedRecord(List<string> fields, CsvLoadOptions options, int lineNumber, bool quotedRecord, Dictionary<string, string>? stringCache)
    {
        try
        {
            PrepareParsedRecord(fields, options, lineNumber, quotedRecord, stringCache);
            return true;
        }
        catch (CsvParseException ex) when (HandleParseError(options, ex, lineNumber))
        {
            return false;
        }
    }

    private static string PrepareParsedField(string value, CsvLoadOptions options, int lineNumber, bool quotedField, Dictionary<string, string>? stringCache)
    {
        if (options.MaxFieldLength is { } maxFieldLength && value.Length > maxFieldLength)
        {
            throw new CsvParseException($"CSV field length {value.Length} exceeds the configured maximum of {maxFieldLength}.", lineNumber);
        }

        if (quotedField && options.MaxQuotedFieldLength is { } maxQuotedFieldLength && value.Length > maxQuotedFieldLength)
        {
            throw new CsvParseException($"CSV quoted field length {value.Length} exceeds the configured maximum of {maxQuotedFieldLength}.", lineNumber);
        }

        if (options.NormalizeQuotes)
        {
            value = NormalizeSmartQuotes(value);
        }

        if (stringCache is not null)
        {
            if (stringCache.TryGetValue(value, out var cached))
            {
                return cached;
            }

            stringCache[value] = value;
        }

        return value;
    }

    private static string NormalizeSmartQuotes(string value)
    {
        var replacementIndex = value.IndexOfAny(new[] { '\u2018', '\u2019', '\u201A', '\u201B', '\u201C', '\u201D', '\u201E', '\u201F' });
        if (replacementIndex < 0)
        {
            return value;
        }

        var chars = value.ToCharArray();
        for (var i = replacementIndex; i < chars.Length; i++)
        {
            chars[i] = chars[i] switch
            {
                '\u2018' or '\u2019' or '\u201A' or '\u201B' => '\'',
                '\u201C' or '\u201D' or '\u201E' or '\u201F' => '"',
                _ => chars[i]
            };
        }

        return new string(chars);
    }
}
