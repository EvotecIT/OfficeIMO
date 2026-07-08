#nullable enable

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal static void ReadRecordsReusableWithMetadataUntilAccepted(
        TextReader reader,
        CsvLoadOptions options,
        Func<CsvParsedRecord, bool> metadataRecordAction,
        Action<IReadOnlyList<string>> recordAction)
    {
        if (metadataRecordAction == null)
        {
            throw new ArgumentNullException(nameof(metadataRecordAction));
        }

        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        ReadLineOrQuotedReusableWithMetadataUntilAccepted(reader, options, metadataRecordAction, recordAction);
    }

    private static void ReadLineOrQuotedReusableWithMetadataUntilAccepted(
        TextReader reader,
        CsvLoadOptions options,
        Func<CsvParsedRecord, bool> metadataRecordAction,
        Action<IReadOnlyList<string>> recordAction)
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
        var metadataAccepted = false;
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
                    PrepareParsedRecord(reusableRecord, options, lineNumber, quotedRecord: false, stringCache);
                    ProcessReusableRecord(reusableRecord, startsWithCommentCharacter: false, metadataRecordAction, recordAction, ref metadataAccepted);
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
            if (!metadataAccepted &&
                TrySkipCommentRecordBeforeParsing(lineReader, pendingLines, startsWithCommentCharacter, line, lineSeparator, options, emittedRecordCount, ref lineNumber))
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
                    ProcessReusableRecord(reusableRecord, startsWithCommentCharacter, metadataRecordAction, recordAction, ref metadataAccepted);
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
                    var logicalRecord = new System.Text.StringBuilder(line);
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

            if (ShouldEmitRecord(reusableQuotedRecord, allowEmpty) &&
                !ShouldSkipCommentRecord(startsWithCommentCharacter, line, options, emittedRecordCount))
            {
                PrepareParsedRecord(reusableQuotedRecord, options, lineNumber, quotedRecord: true, stringCache);
                ProcessReusableRecord(reusableQuotedRecord, startsWithCommentCharacter, metadataRecordAction, recordAction, ref metadataAccepted);
                emittedRecordCount++;
                ReportProgress(options, emittedRecordCount, lineNumber);
            }

            lineNumber++;
        }
    }

    private static void ProcessReusableRecord(
        IReadOnlyList<string> record,
        bool startsWithCommentCharacter,
        Func<CsvParsedRecord, bool> metadataRecordAction,
        Action<IReadOnlyList<string>> recordAction,
        ref bool metadataAccepted)
    {
        if (metadataAccepted)
        {
            recordAction(record);
            return;
        }

        metadataAccepted = metadataRecordAction(new CsvParsedRecord(record, startsWithCommentCharacter));
    }
}
