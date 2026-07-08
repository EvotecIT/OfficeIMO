#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal static IEnumerable<IReadOnlyList<string>> ParseReusable(TextReader reader, CsvLoadOptions options)
    {
        return ParseLineOrQuotedReusable(reader, options);
    }

    private static IEnumerable<IReadOnlyList<string>> ParseLineOrQuotedReusable(TextReader reader, CsvLoadOptions options)
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
        var stringCache = CreateStringCache(options);
        using var lineReader = new CsvLineReader(reader);

        while (true)
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
                    PrepareParsedRecord(reusableRecord, options, lineNumber, quotedRecord: false, stringCache);
                    yield return reusableRecord;
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
                    PrepareParsedRecord(reusableRecord, options, lineNumber, quotedRecord: false, stringCache);
                    yield return reusableRecord;
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }

                lineNumber++;
                continue;
            }

            try
            {
                if (!TryParseQuotedRecord(line, delimiter, trim, strictQuotes, lineNumber, reusableQuotedRecord))
                {
                    var logicalRecord = new StringBuilder(line);
                    var pendingSeparator = lineSeparator;
                    while (true)
                    {
                        ThrowIfCancellationRequested(options);
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
                    PrepareParsedRecord(reusableQuotedRecord, options, lineNumber, quotedRecord: true, stringCache);
                    yield return reusableQuotedRecord;
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }
            }

            lineNumber++;
        }
    }
}
