#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal sealed partial class CsvUtf8StreamDataReaderRowSource
    {
        /// <summary>
        /// Skips a comment record whose first line contains a quote, as canonical comment skipping
        /// does: when the line leaves a quote open, following lines join the comment until the quote
        /// closes, at most <see cref="CsvParsingSettings.MaxCommentContinuationLines"/> of them. Lines
        /// read ahead are put back when the quote stays open, at the end of input, or when a
        /// delimited line follows an undelimited comment. Returns false when the record cannot be
        /// buffered.
        /// </summary>
        private bool SkipQuotedCommentRecord(ref int recordStart, CancellationToken cancellationToken)
        {
            while (true)
            {
                if (TryScanQuotedComment(recordStart, out int next, out int separators))
                {
                    _position = next;
                    _physicalLineSeparatorsConsumed += separators;
                    return true;
                }

                if (!FillForCurrentRecord(ref recordStart, cancellationToken) && !_endOfStream)
                {
                    _position = recordStart;
                    return false;
                }
            }
        }

        /// <summary>Finds the extent of a quoted comment record; false when more input is needed.</summary>
        private bool TryScanQuotedComment(int recordStart, out int next, out int separators)
        {
            next = recordStart;
            separators = 0;
            if (!TryFindLineEnd(recordStart, out int firstEnd, out int firstNext))
            {
                return false;
            }

            ReadOnlySpan<byte> delimiter = _delimiterByte;
            next = firstNext;
            separators = firstNext > firstEnd ? 1 : 0;
            var state = new TextQuoteState();
            if (ParseTextDelimitedQuotes(recordStart, firstEnd, delimiter, strict: false, visit: false, ref state) != TextQuoteResult.InQuotes)
            {
                return true;
            }

            int limit = _options.MaxCommentContinuationLines;
            if (limit <= 0)
            {
                throw new ArgumentOutOfRangeException("options", "MaxCommentContinuationLines must be greater than zero.");
            }

            bool firstLineLooksDelimited = LooksLikeDelimitedComment(_buffer.AsSpan(recordStart, firstEnd - recordStart), delimiter);
            int position = firstNext;
            int lineSeparators = separators;
            for (int continuations = 0; continuations < limit; continuations++)
            {
                if (position == _length)
                {
                    return _endOfStream;
                }

                if (!TryFindLineEnd(position, out int lineEnd, out int lineNext))
                {
                    return false;
                }

                ParseTextDelimitedQuotes(position, lineEnd, delimiter, strict: false, visit: false, ref state);
                lineSeparators += lineNext > lineEnd ? 1 : 0;
                if (!state.InQuotes)
                {
                    next = lineNext;
                    separators = lineSeparators;
                    return true;
                }

                if (!firstLineLooksDelimited && LooksLikeDelimitedComment(_buffer.AsSpan(position, lineEnd - position), delimiter))
                {
                    return true;
                }

                position = lineNext;
            }

            return true;
        }

        private static bool LooksLikeDelimitedComment(ReadOnlySpan<byte> line, ReadOnlySpan<byte> delimiter) =>
            line.IndexOf(delimiter) >= 0 || line.IndexOfAny(DelimitedCommentBytes) >= 0;
    }
}
#endif
