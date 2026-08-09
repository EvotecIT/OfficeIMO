#nullable enable

using System.Buffers;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    internal static void ReadFieldSpans<TVisitor>(
        string text,
        CsvLoadOptions options,
        int recordsToSkip,
        ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (!TryReadFieldSpansWithTextDataReaderBatch(text, options, recordsToSkip, ref fieldVisitor))
        {
            ReadFieldSpans(text.AsSpan(), options, recordsToSkip, ref fieldVisitor);
        }
    }

    private static bool TryReadFieldSpansWithTextDataReaderBatch<TVisitor>(
        string text,
        CsvLoadOptions options,
        int recordsToSkip,
        ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (HasFieldLengthLimits(options) ||
            UsesTextDelimiter(options) ||
            (NeedsLogicalCommentSkipping(options) &&
                HasPotentialTextCommentRecord(text.AsSpan(), options.CommentCharacter)) ||
            options.ParseErrorAction == CsvParseErrorAction.SkipRow ||
            !CanUseTextDataReaderBatchAvx2(options, TextQuoteAwareFieldSpanCapacity))
        {
            return false;
        }

        var state = CreateTextFieldSpanReadState(text.AsSpan(), options, recordsToSkip);
        var projectedFieldVisitor = fieldVisitor as ICsvProjectedFieldSpanVisitor;
        var batch = CreateFieldSpanTextDataReaderBatch(
            text,
            options,
            sourceColumnCount: 16);
        try
        {
            SkipPendingTextDataReaderRecords(
                text.AsSpan(),
                options,
                options.CancellationToken,
                ref state);

            while (state.Position < text.Length)
            {
                if (TryFillTextDataReaderBatchAvx2(
                    text.AsSpan(),
                    options,
                    ref state,
                    batch,
                    options.CancellationToken))
                {
                    var recordIndex = state.RecordIndex - batch.RowCount;
                    while (batch.MoveNext())
                    {
                        ThrowIfCancellationRequested(options);
                        batch.VisitCurrentRow(
                            recordIndex++,
                            projectedFieldVisitor,
                            ref fieldVisitor,
                            ref state.Scratch);
                    }
                    continue;
                }

                if (!TryReadNextTextRecordFieldSpans(
                    text.AsSpan(),
                    options,
                    projectedFieldVisitor,
                    ref state,
                    ref fieldVisitor,
                    out var emittedFieldCount))
                {
                    break;
                }

                if (emittedFieldCount > batch.SourceColumnCount &&
                    emittedFieldCount <= TextQuoteAwareFieldSpanCapacity)
                {
                    var replacement = CreateFieldSpanTextDataReaderBatch(
                        text,
                        options,
                        GetTextDelimiterIndexCapacity(emittedFieldCount));
                    batch.Dispose();
                    batch = replacement;
                }
            }
        }
        finally
        {
            batch.Dispose();
            if (state.Scratch is not null)
            {
                ArrayPool<char>.Shared.Return(state.Scratch);
            }
        }

        return true;
    }

    private static CsvTextDataReaderBatch CreateFieldSpanTextDataReaderBatch(
        string text,
        CsvLoadOptions options,
        int sourceColumnCount) =>
        new(
            text,
            sourceColumnCount,
            options.Culture,
            options.DateTimeFormats,
            preferredRowCapacity: 512,
            rejectExtraFields: true,
            enforceColumnCountMismatchPolicy: false);
#endif
}
