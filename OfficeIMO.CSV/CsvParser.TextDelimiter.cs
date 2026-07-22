#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    private static bool UsesTextDelimiter(CsvLoadOptions options) =>
        !string.IsNullOrEmpty(options.DelimiterText) && options.DelimiterText!.Length > 1;

    private static char GetDelimiterChar(CsvLoadOptions options) =>
        string.IsNullOrEmpty(options.DelimiterText)
            ? options.Delimiter
            : options.DelimiterText![0];

    private static string GetDelimiterText(CsvLoadOptions options) =>
        string.IsNullOrEmpty(options.DelimiterText)
            ? options.Delimiter.ToString()
            : options.DelimiterText!;

    private static IEnumerable<string[]> ParseLineOrQuotedTextDelimiter(TextReader reader, CsvLoadOptions options)
    {
        foreach (var record in ParseLineOrQuotedTextDelimiterWithMetadata(reader, options))
        {
            if (record.Values is string[] fields)
            {
                yield return fields;
            }
            else
            {
                yield return record.Values.ToArray();
            }
        }
    }

    private static void ReadLineOrQuotedTextDelimiter(TextReader reader, CsvLoadOptions options, Action<string[]> recordAction)
    {
        foreach (var record in ParseLineOrQuotedTextDelimiter(reader, options))
        {
            recordAction(record);
        }
    }

    private static void ReadLineOrQuotedTextDelimiterReusable(TextReader reader, CsvLoadOptions options, Action<IReadOnlyList<string>> recordAction)
    {
        foreach (var record in ParseLineOrQuotedTextDelimiterWithMetadata(reader, options))
        {
            recordAction(record.Values);
        }
    }

    private static void ReadLineOrQuotedTextDelimiterReusableWithMetadata(TextReader reader, CsvLoadOptions options, Action<CsvParsedRecord> recordAction)
    {
        foreach (var record in ParseLineOrQuotedTextDelimiterWithMetadata(reader, options))
        {
            recordAction(record);
        }
    }

    private static IEnumerable<CsvParsedRecord> ParseLineOrQuotedTextDelimiterWithMetadata(TextReader reader, CsvLoadOptions options)
    {
        var delimiter = GetDelimiterText(options);
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var lineNumber = 1;
        var emittedRecordCount = 0;
        var stringCache = CreateStringCache(options);
        using var lineReader = new CsvLineReader(reader);

        while (lineReader.ReadLine(out var lineSeparator) is { } line)
        {
            ThrowIfCancellationRequested(options);
            var startsWithCommentCharacter = IsRawCommentLine(line, options);
            if (TrySkipTextDelimiterCommentRecordBeforeParsing(
                    lineReader,
                    startsWithCommentCharacter,
                    line,
                    lineSeparator,
                    options,
                    emittedRecordCount,
                    delimiter,
                    trim,
                    ref lineNumber))
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
                if (!TryParseQuotedRecord(line, delimiter, trim, strictQuotes, lineNumber, out fields))
                {
                    var logicalRecord = new StringBuilder(line);
                    var pendingSeparator = lineSeparator;
                    var inQuotes = false;
                    UpdateQuotedRecordState(line, ref inQuotes);
                    while (inQuotes)
                    {
                        ThrowIfCancellationRequested(options);
                        var next = lineReader.ReadLine(out var nextSeparator);
                        if (next == null)
                        {
                            throw new CsvParseException("Unterminated quoted field.", lineNumber);
                        }

                        logicalRecord.Append(pendingSeparator);
                        logicalRecord.Append(next);
                        lineNumber++;
                        UpdateQuotedRecordState(next, ref inQuotes);
                        pendingSeparator = nextSeparator;
                    }

                    if (!TryParseQuotedRecord(logicalRecord.ToString(), delimiter, trim, strictQuotes, lineNumber, out fields))
                    {
                        throw new CsvParseException("Unterminated quoted field.", lineNumber);
                    }
                }
            }
            catch (CsvParseException ex) when (HandleParseError(options, ex, lineNumber))
            {
                lineNumber++;
                continue;
            }

            if (ShouldEmitRecord(fields, allowEmpty) &&
                !ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
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

            lineNumber++;
        }
    }

    private static bool TrySplitUnquotedRecord(string line, string delimiter, bool trim, out string[] fields)
    {
        fields = Array.Empty<string>();
        if (line.IndexOf('"') >= 0)
        {
            return false;
        }

        var parsedFields = new List<string>();
        var start = 0;
        while (true)
        {
            var delimiterIndex = line.IndexOf(delimiter, start, StringComparison.Ordinal);
            if (delimiterIndex < 0)
            {
                parsedFields.Add(GetUnquotedField(line, start, line.Length - start, trim));
                fields = parsedFields.ToArray();
                return true;
            }

            parsedFields.Add(GetUnquotedField(line, start, delimiterIndex - start, trim));
            start = delimiterIndex + delimiter.Length;
        }
    }

    private static bool TryParseQuotedRecord(string text, string delimiter, bool trim, bool strictQuotes, int lineNumber, out string[] fields)
    {
        var parsedFields = new List<string>(16);
        var buffer = new StringBuilder();
        var inQuotes = false;
        var fieldWasQuoted = false;
        var afterClosingQuote = false;
        var bufferIsWhitespaceOnly = true;

        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        i++;
                        buffer.Append('"');
                        bufferIsWhitespaceOnly = false;
                    }
                    else
                    {
                        inQuotes = false;
                        afterClosingQuote = true;
                    }
                }
                else
                {
                    buffer.Append(c);
                    if (!char.IsWhiteSpace(c))
                    {
                        bufferIsWhitespaceOnly = false;
                    }
                }

                continue;
            }

            if (StartsWithDelimiter(text, delimiter, i))
            {
                AddQuotedField(parsedFields, buffer, trim, ref fieldWasQuoted);
                bufferIsWhitespaceOnly = true;
                afterClosingQuote = false;
                i += delimiter.Length - 1;
                continue;
            }

            if (c == '"')
            {
                if (afterClosingQuote)
                {
                    if (strictQuotes)
                    {
                        throw new CsvParseException("Invalid quoted field.", lineNumber);
                    }

                    buffer.Append(c);
                    bufferIsWhitespaceOnly = false;
                    afterClosingQuote = false;
                    continue;
                }

                if (trim && bufferIsWhitespaceOnly)
                {
                    buffer.Clear();
                }

                inQuotes = true;
                fieldWasQuoted = true;
                continue;
            }

            if (afterClosingQuote && char.IsWhiteSpace(c) && trim)
            {
                continue;
            }

            if (afterClosingQuote && strictQuotes)
            {
                throw new CsvParseException("Invalid quoted field.", lineNumber);
            }

            afterClosingQuote = false;
            buffer.Append(c);
            if (!char.IsWhiteSpace(c))
            {
                bufferIsWhitespaceOnly = false;
            }
        }

        if (inQuotes)
        {
            fields = Array.Empty<string>();
            return false;
        }

        AddQuotedField(parsedFields, buffer, trim, ref fieldWasQuoted);
        fields = parsedFields.ToArray();
        return true;
    }

    private static bool StartsWithDelimiter(string text, string delimiter, int index) =>
        index <= text.Length - delimiter.Length &&
        string.CompareOrdinal(text, index, delimiter, 0, delimiter.Length) == 0;

    private static bool TrySkipTextDelimiterCommentRecordBeforeParsing(
        CsvLineReader reader,
        bool startsWithCommentCharacter,
        string firstLine,
        string firstLineSeparator,
        CsvLoadOptions options,
        int emittedRecordCount,
        string delimiter,
        bool trim,
        ref int lineNumber)
    {
        if (!ShouldSkipCommentRecordBeforeParsing(startsWithCommentCharacter, firstLine, options, emittedRecordCount))
        {
            return false;
        }

        if (firstLine.IndexOf('"') < 0 ||
            !LooksLikeDelimitedRawComment(firstLine, delimiter) ||
            TryParseQuotedRecord(firstLine, delimiter, trim, strictQuotes: false, lineNumber, out _))
        {
            return true;
        }

        var continuationCount = 0;
        var inQuotes = false;
        UpdateQuotedRecordState(firstLine, ref inQuotes);
        while (true)
        {
            ThrowIfCancellationRequested(options);
            var next = reader.ReadLine(out _);
            if (next == null)
            {
                return true;
            }

            continuationCount++;
            UpdateQuotedRecordState(next, ref inQuotes);
            if (!inQuotes)
            {
                lineNumber += continuationCount;
                return true;
            }
        }
    }

    private static bool LooksLikeDelimitedRawComment(string line, string delimiter) =>
        line.IndexOf(delimiter, StringComparison.Ordinal) >= 0 ||
        line.IndexOf(',') >= 0 ||
        line.IndexOf(';') >= 0 ||
        line.IndexOf('|') >= 0 ||
        line.IndexOf('\t') >= 0;

#if NET8_0_OR_GREATER
    private static void ReadFieldSpansTextDelimiter<TVisitor>(TextReader reader, CsvLoadOptions options, int recordsToSkip, ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (recordsToSkip < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(recordsToSkip), "SkipInitialRecords cannot be negative.");
        }

        var recordIndex = 0;
        var projectedFieldVisitor = fieldVisitor as ICsvProjectedFieldSpanVisitor;
        foreach (var record in ParseLineOrQuotedTextDelimiterWithMetadata(reader, options))
        {
            if (recordsToSkip > 0)
            {
                recordsToSkip--;
                continue;
            }

            VisitParsedFields(record.Values, recordIndex, projectedFieldVisitor, ref fieldVisitor);
            recordIndex++;
        }
    }
#endif
}
