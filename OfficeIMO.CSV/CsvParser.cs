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
        return ParseLineOrQuoted(reader, options);
    }

    public static void ReadRecords(TextReader reader, CsvLoadOptions options, Action<string[]> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        ReadLineOrQuoted(reader, options, recordAction);
    }

    internal static IEnumerable<CsvParsedRecord> ParseWithMetadata(TextReader reader, CsvLoadOptions options)
    {
        return ParseLineOrQuotedWithMetadata(reader, options);
    }

    internal static void ReadRecordsWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        foreach (var record in ParseLineOrQuotedWithMetadata(reader, options))
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

        ReadLineOrQuotedReusable(reader, options, recordAction);
    }

    internal static void ReadRecordsReusableWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
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

        ReadFieldSpansLineOrQuoted(reader, options, recordsToSkip, ref fieldVisitor);
    }
#endif

    private static void ReadLineOrQuoted(TextReader reader, CsvLoadOptions options, Action<string[]> recordAction)
    {
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();
        using var lineReader = new CsvLineReader(reader);

        while (ReadLineWithSeparator(lineReader, pendingLines, out var lineSeparator) is { } line)
        {
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
                    recordAction(record);
                    emittedRecordCount++;
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            if (!TryParseQuotedRecord(line, delimiter, trim, strictQuotes, lineNumber, out fields))
            {
                var logicalRecord = new StringBuilder(line);
                var pendingSeparator = lineSeparator;
                while (true)
                {
                    var next = ReadLineWithSeparator(lineReader, pendingLines, out var nextSeparator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(pendingSeparator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, strictQuotes, lineNumber, out fields))
                    {
                        break;
                    }

                    pendingSeparator = nextSeparator;
                }
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    recordAction(fields);
                    emittedRecordCount++;
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
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();
        using var lineReader = new CsvLineReader(reader);

        while (ReadLineWithSeparator(lineReader, pendingLines, out var lineSeparator) is { } line)
        {
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
                    yield return new CsvParsedRecord(record, startsWithCommentCharacter);
                    emittedRecordCount++;
                }

                lineNumber++;
                continue;
            }

            string[] fields;
            if (!TryParseQuotedRecord(line, delimiter, trim, strictQuotes, lineNumber, out fields))
            {
                var logicalRecord = new StringBuilder(line);
                var pendingSeparator = lineSeparator;
                while (true)
                {
                    var next = ReadLineWithSeparator(lineReader, pendingLines, out var nextSeparator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(pendingSeparator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, strictQuotes, lineNumber, out fields))
                    {
                        break;
                    }

                    pendingSeparator = nextSeparator;
                }
            }

            if (ShouldEmitRecord(fields, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    yield return new CsvParsedRecord(fields, startsWithCommentCharacter);
                    emittedRecordCount++;
                }
            }

            lineNumber++;
        }
    }

    private static void ReadLineOrQuotedReusable(TextReader reader, CsvLoadOptions options, Action<IReadOnlyList<string>> recordAction)
    {
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var reusableRecord = new List<string>(64);
        var reusableQuotedRecord = new List<string>(64);
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();
        using var lineReader = new CsvLineReader(reader);

        while (pendingLines.Count > 0 || true)
        {
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
                    recordAction(reusableRecord);
                    emittedRecordCount++;
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
                    recordAction(reusableRecord);
                    emittedRecordCount++;
                }

                lineNumber++;
                continue;
            }

            if (!TryParseQuotedRecord(line, delimiter, trim, strictQuotes, lineNumber, reusableQuotedRecord))
            {
                var logicalRecord = new StringBuilder(line);
                var pendingSeparator = lineSeparator;
                while (true)
                {
                    var next = ReadLineWithSeparator(lineReader, pendingLines, out var nextSeparator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(pendingSeparator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, strictQuotes, lineNumber, reusableQuotedRecord))
                    {
                        break;
                    }

                    pendingSeparator = nextSeparator;
                }
            }

            if (ShouldEmitRecord(reusableQuotedRecord, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    recordAction(reusableQuotedRecord);
                    emittedRecordCount++;
                }
            }

            lineNumber++;
        }
    }

    private static void ReadLineOrQuotedReusableWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var reusableRecord = new List<string>(64);
        var reusableQuotedRecord = new List<string>(64);
        var emittedRecordCount = 0;
        var pendingLines = new Queue<CsvLine>();
        using var lineReader = new CsvLineReader(reader);

        while (pendingLines.Count > 0 || true)
        {
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
                    recordAction(new CsvParsedRecord(reusableRecord, startsWithCommentCharacter: false));
                    emittedRecordCount++;
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
                    recordAction(new CsvParsedRecord(reusableRecord, startsWithCommentCharacter));
                    emittedRecordCount++;
                }

                lineNumber++;
                continue;
            }

            if (!TryParseQuotedRecord(line, delimiter, trim, strictQuotes, lineNumber, reusableQuotedRecord))
            {
                var logicalRecord = new StringBuilder(line);
                var pendingSeparator = lineSeparator;
                while (true)
                {
                    var next = ReadLineWithSeparator(lineReader, pendingLines, out var nextSeparator);
                    if (next == null)
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }

                    logicalRecord.Append(pendingSeparator);
                    logicalRecord.Append(next);
                    lineNumber++;

                    if (TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, strictQuotes, lineNumber, reusableQuotedRecord))
                    {
                        break;
                    }

                    pendingSeparator = nextSeparator;
                }
            }

            if (ShouldEmitRecord(reusableQuotedRecord, allowEmpty))
            {
                if (!ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
                {
                    recordAction(new CsvParsedRecord(reusableQuotedRecord, startsWithCommentCharacter));
                    emittedRecordCount++;
                }
            }

            lineNumber++;
        }
    }

#if NET8_0_OR_GREATER
    private static void ReadFieldSpansLineOrQuoted<TVisitor>(TextReader reader, CsvLoadOptions options, int recordsToSkip, ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var delimiter = options.Delimiter;
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
                }

                lineNumber++;
                continue;
            }

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

        if (TryParseQuotedRecordLenient(firstLine, options.Delimiter, options.TrimWhitespace, out _))
        {
            return true;
        }

        var logicalRecord = new StringBuilder(firstLine);
        var pendingSeparator = firstLineSeparator;
        var continuations = new List<CsvLine>();
        while (true)
        {
            var next = ReadLineWithSeparator(reader, pendingLines, out var nextSeparator);
            if (next == null)
            {
                EnqueueContinuations(pendingLines, continuations);
                return true;
            }

            continuations.Add(new CsvLine(next, nextSeparator));
            var candidate = string.Concat(logicalRecord.ToString(), pendingSeparator, next);
            if (TryParseQuotedRecordLenient(candidate, options.Delimiter, options.TrimWhitespace, out _))
            {
                lineNumber += continuations.Count;
                return true;
            }

            if (!LooksLikeDelimitedRawComment(firstLine, options.Delimiter) && LooksLikeDelimitedRawComment(next, options.Delimiter))
            {
                EnqueueContinuations(pendingLines, continuations);
                return true;
            }

            logicalRecord.Append(pendingSeparator);
            logicalRecord.Append(next);
            pendingSeparator = nextSeparator;
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
}
