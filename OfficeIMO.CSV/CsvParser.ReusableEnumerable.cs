#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal static IEnumerable<IReadOnlyList<string>> ParseReusable(TextReader reader, CsvLoadOptions options)
    {
        if (UsesTextDelimiter(options))
        {
            foreach (var record in ParseLineOrQuotedTextDelimiterWithMetadata(reader, options))
            {
                yield return record.Values;
            }

            yield break;
        }

        foreach (var record in ParseLineOrQuotedReusable(reader, options))
        {
            yield return record;
        }
    }

    private static IEnumerable<IReadOnlyList<string>> ParseLineOrQuotedReusable(TextReader reader, CsvLoadOptions options)
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
                    if (!TryPrepareParsedRecord(reusableRecord, options, lineNumber, quotedRecord: false, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

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
                    if (!TryPrepareParsedRecord(reusableRecord, options, lineNumber, quotedRecord: false, stringCache))
                    {
                        lineNumber++;
                        continue;
                    }

                    yield return reusableRecord;
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

                    yield return reusableQuotedRecord;
                    emittedRecordCount++;
                    ReportProgress(options, emittedRecordCount, lineNumber);
                }
            }

            lineNumber++;
        }
    }
}
