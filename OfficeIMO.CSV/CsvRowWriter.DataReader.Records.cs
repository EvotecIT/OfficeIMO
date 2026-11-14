using System.Data;
using System.Text;
using OfficeIMO.SharedSource.IO;

namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    /// <summary>
    /// Writes typed data-reader rows through a pooled buffer: common values are formatted straight into it instead
    /// of through a row string builder, and completed rows reach the writer in batches. Rows of comma-separated
    /// invariant values go to a UTF-8 stream as bytes. Only completed rows are written, so failures and cancellation
    /// leave the same output as before.
    /// </summary>
    private void WriteDataReaderRecords(IDataReader reader, CsvWriter.DataReaderFieldKind[] fieldKinds, CancellationToken cancellationToken)
    {
        // Data table rows look values up per call; one boxed read answers both the null check and the value.
        Type readerType = reader.GetType();
        bool boxedValues = readerType == typeof(DataTableReader) || readerType == typeof(CsvDataTableRowReader);
        bool rowPending = false;
#if NET8_0_OR_GREATER
        CsvWriter.InvariantRecordLayout layout = !_useFormattedDefaultWritePath && _delimiter == ','
            && ReferenceEquals(_options.Culture, CultureInfo.InvariantCulture)
                ? CsvWriter.GetInvariantRecordLayout(fieldKinds)
                : CsvWriter.InvariantRecordLayout.None;
        bool invariantRows = layout != CsvWriter.InvariantRecordLayout.None;
        // Exactly UTF8Encoding: its encoder writes ASCII as is, so the bytes match what the text path would encode.
        if (invariantRows && _writer is PooledUtf8TextWriter utf8Writer && utf8Writer.Encoding.GetType() == typeof(UTF8Encoding))
        {
            rowPending = WriteInvariantUtf8Records(reader, utf8Writer, fieldKinds, layout, boxedValues, cancellationToken);
            if (!rowPending)
            {
                return;
            }

            invariantRows = false;
        }
#endif

        var scratch = new StringBuilder();
        using var output = new CsvWriter.RecordBuffer<char>();
        var completed = 0;
        try
        {
            while (true)
            {
                if (!rowPending)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!reader.Read())
                    {
                        break;
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                }

                rowPending = false;
#if NET8_0_OR_GREATER
                if (!invariantRows || !TryAppendInvariantRecord(output, reader, fieldKinds, layout, boxedValues, _options.NewLine.AsSpan(), null))
#endif
                {
                    // After a rejected typed read every later row takes the per-field path, as before.
#if NET8_0_OR_GREATER
                    invariantRows = false;
#endif
                    CsvWriter.AppendDataReaderRecord(output, scratch, reader, fieldKinds, _delimiter, _options.NewLine,
                        _options.Culture, _useFormattedDefaultWritePath, _options.DateTimeFormat, _options.UseUtc, _options.NullValue, boxedValues);
                }

                completed = output.Length;
                if (completed >= CsvWriter.DataReaderFlushThreshold)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Deliver(_writer, output, ref completed);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            Deliver(_writer, output, ref completed);
        }
        catch
        {
            Deliver(_writer, output, ref completed);
            throw;
        }
    }

    /// <summary>Writes the completed rows and empties the buffer, so a failing write is not repeated.</summary>
    private static void Deliver(TextWriter writer, CsvWriter.RecordBuffer<char> output, ref int completed)
    {
        int length = completed;
        completed = 0;
        output.Length = 0;
        if (length != 0)
        {
#if NET8_0_OR_GREATER
            writer.Write(output.Prefix(length));
#else
            writer.Write(output.Units, 0, length);
#endif
        }
    }
#if NET8_0_OR_GREATER

    /// <summary>
    /// Writes rows of invariant values to a UTF-8 stream writer as bytes. Returns true when a typed read rejected
    /// the current row, after delivering the rows before it; the caller then writes that row and the rest as text.
    /// </summary>
    private bool WriteInvariantUtf8Records(
        IDataReader reader, PooledUtf8TextWriter writer, CsvWriter.DataReaderFieldKind[] fieldKinds, CsvWriter.InvariantRecordLayout layout,
        bool boxedValues, CancellationToken cancellationToken)
    {
        Encoding encoding = writer.Encoding;
        byte[] newLine = encoding.GetBytes(_options.NewLine);
        using var output = new CsvWriter.RecordBuffer<byte>();
        var completed = 0;
        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!reader.Read())
                {
                    break;
                }

                cancellationToken.ThrowIfCancellationRequested();
                if (!TryAppendInvariantRecord(output, reader, fieldKinds, layout, boxedValues, newLine, encoding))
                {
                    Deliver(writer, output, ref completed);
                    return true;
                }

                completed = output.Length;
                if (completed >= CsvWriter.DataReaderFlushThreshold)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Deliver(writer, output, ref completed);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            Deliver(writer, output, ref completed);
            return false;
        }
        catch
        {
            Deliver(writer, output, ref completed);
            throw;
        }
    }

    /// <inheritdoc cref="Deliver(TextWriter, CsvWriter.RecordBuffer{char}, ref int)"/>
    private static void Deliver(PooledUtf8TextWriter writer, CsvWriter.RecordBuffer<byte> output, ref int completed)
    {
        int length = completed;
        completed = 0;
        output.Length = 0;
        if (length != 0)
        {
            writer.WriteUtf8(output.Prefix(length));
        }
    }

    /// <summary>Writes one row through the one-pass invariant path; false rewinds it for the per-field path.</summary>
    private static bool TryAppendInvariantRecord<TChar>(
        CsvWriter.RecordBuffer<TChar> output, IDataReader reader, CsvWriter.DataReaderFieldKind[] fieldKinds,
        CsvWriter.InvariantRecordLayout layout, bool boxedValues, ReadOnlySpan<TChar> newLine, Encoding? encoding)
        where TChar : unmanaged
    {
        int rowStart = output.Length;
        try
        {
            if (boxedValues)
            {
                var fields = new CsvWriter.BoxedRecordFields(reader);
                AppendInvariantRecord(output, ref fields, fieldKinds, layout, newLine, encoding);
            }
            else
            {
                var fields = new CsvWriter.TypedRecordFields(reader);
                AppendInvariantRecord(output, ref fields, fieldKinds, layout, newLine, encoding);
            }

            return true;
        }
        catch (Exception exception) when (exception is InvalidCastException or NotSupportedException or NotImplementedException)
        {
            output.Length = rowStart;
            return false;
        }
    }

    private static void AppendInvariantRecord<TFields, TChar>(
        CsvWriter.RecordBuffer<TChar> output, ref TFields fields, CsvWriter.DataReaderFieldKind[] fieldKinds,
        CsvWriter.InvariantRecordLayout layout, ReadOnlySpan<TChar> newLine, Encoding? encoding)
        where TFields : struct, CsvWriter.IRecordFields
        where TChar : unmanaged
    {
        switch (layout)
        {
            case CsvWriter.InvariantRecordLayout.Text:
                CsvWriter.AppendInvariantRecord<TFields, TChar, CsvWriter.TextRecordLayout>(output, ref fields, fieldKinds, newLine, encoding);
                break;
            case CsvWriter.InvariantRecordLayout.Scalar:
                CsvWriter.AppendInvariantRecord<TFields, TChar, CsvWriter.ScalarRecordLayout>(output, ref fields, fieldKinds, newLine, encoding);
                break;
            default:
                CsvWriter.AppendInvariantRecord<TFields, TChar, CsvWriter.MixedRecordLayout>(output, ref fields, fieldKinds, newLine, encoding);
                break;
        }
    }
#endif
}
