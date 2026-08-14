#nullable enable

using System.Buffers;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private static bool TryFillTextDataReaderBatchScalar(
        string text,
        CsvLoadOptions options,
        ref CsvTextFieldSpanReadState state,
        CsvTextDataReaderBatch batch,
        CancellationToken cancellationToken)
    {
        batch.Reset();
        while (state.Position < text.Length && !batch.IsFull)
        {
            cancellationToken.ThrowIfCancellationRequested();
            CsvTextFieldSpanReadState rowState = state;
            var visitor = new CsvTextDataReaderBatchVisitor(text, batch);
            if (!TryReadNextTextRecordFieldSpans(
                    text.AsSpan(),
                    options,
                    null,
                    ref state,
                    ref visitor,
                    out int fieldCount))
            {
                break;
            }

            if (visitor.Failed)
            {
                if (!ReferenceEquals(state.Scratch, rowState.Scratch) && state.Scratch is not null)
                {
                    ArrayPool<char>.Shared.Return(state.Scratch);
                }
                state = rowState;
                batch.DiscardPendingRow();
                return batch.RowCount != 0;
            }

            batch.CompleteRow(fieldCount, options.ColumnCountMismatchPolicy);
        }

        return batch.RowCount != 0;
    }

    private struct CsvTextDataReaderBatchVisitor : ICsvFieldSpanVisitor
    {
        private readonly string _text;
        private readonly CsvTextDataReaderBatch _batch;

        internal CsvTextDataReaderBatchVisitor(string text, CsvTextDataReaderBatch batch)
        {
            _text = text;
            _batch = batch;
            Failed = false;
        }

        internal bool Failed { get; private set; }

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            if (Failed)
            {
                return;
            }

            ref char textStart = ref MemoryMarshal.GetReference(_text.AsSpan());
            ref char valueStart = ref MemoryMarshal.GetReference(value);
            nint byteOffset = Unsafe.ByteOffset(ref textStart, ref valueStart);
            if (byteOffset < 0 || (byteOffset & 1) != 0)
            {
                Failed = true;
                return;
            }

            int start = checked((int)(byteOffset / 2));
            Failed = !_batch.TrySetPendingField(fieldIndex, start, value.Length, escapedSourceLength: 0);
        }

        public void VisitFieldRange(int recordIndex, int fieldIndex, char[] buffer, int start, int length)
        {
            Failed = true;
        }

        public bool TryVisitEscapedField(
            int recordIndex,
            int fieldIndex,
            ReadOnlySpan<char> escapedValue,
            int unescapedLength)
        {
            if (Failed)
            {
                return true;
            }

            ref char textStart = ref MemoryMarshal.GetReference(_text.AsSpan());
            ref char valueStart = ref MemoryMarshal.GetReference(escapedValue);
            nint byteOffset = Unsafe.ByteOffset(ref textStart, ref valueStart);
            if (byteOffset < 0 || (byteOffset & 1) != 0)
            {
                Failed = true;
                return true;
            }

            int start = checked((int)(byteOffset / 2));
            Failed = !_batch.TrySetPendingField(
                fieldIndex,
                start,
                unescapedLength,
                escapedValue.Length);
            return true;
        }

        public void VisitFieldValue(int recordIndex, int fieldIndex, string value)
        {
            Failed = true;
        }
    }
#endif
}
