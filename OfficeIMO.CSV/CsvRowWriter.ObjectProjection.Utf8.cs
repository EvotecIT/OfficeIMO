#if NET8_0_OR_GREATER
using System.Buffers;
using OfficeIMO.SharedSource.IO;

namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    /// <summary>Where projected object rows are staged before delivery.</summary>
    internal enum ProjectedStaging : byte
    {
        /// <summary>The row string builder, for writers without a span path.</summary>
        Builder,

        /// <summary>A bounded UTF-16 row buffer.</summary>
        Characters,

        /// <summary>A bounded UTF-8 row buffer for UTF-8 stream writers.</summary>
        Utf8
    }

    /// <summary>Units of the bounded row buffer; longer rows finish through the row string builder.</summary>
    private const int PropertyStagingLength = 1024;

    private PooledUtf8TextWriter? _propertyUtf8Writer;
    private byte[]? _propertyNewLineBytes;

    /// <summary>
    /// The UTF-8 writer's byte buffer while a row is staged in it, from <see cref="_propertyOrigin"/> up to
    /// <see cref="_propertyLimit"/>; <see cref="_propertyCharacterCount"/> is then the absolute end of the staged bytes.
    /// </summary>
    private byte[]? _propertyBytes;
    private int _propertyOrigin;
    private int _propertyLimit;

    /// <summary>
    /// Stages text as UTF-8 bytes. ASCII text is copied or escaped in one pass; other text is escaped and then
    /// encoded by the writer's own encoding, so invalid UTF-16 gets the replacement or the exception it gives.
    /// </summary>
    private void StageProjectedUtf8Text(int ordinal, string? text)
    {
        int start = _propertyCharacterCount;
        int length = text?.Length ?? 0;
        if (_propertyFallbackColumn < 0 && (long)start + length * 2L + 3 <= _propertyLimit)
        {
            int position = start;
            if (ordinal > 0) _propertyBytes![position++] = (byte)',';
            if (text == null)
            {
                _propertyCharacterCount = position;
                return;
            }

            // Projected rows are comma-separated; the room checked above holds every escaped form of ASCII text.
            Span<byte> destination = _propertyBytes.AsSpan(position);
            // Short text is copied plainly first; longer text the vector escape handles skips that copy.
            if ((length <= CsvWriter.ShortTextLength || length > CsvWriter.EscapedShortTextLength) && CsvWriter.TryCopyPlainText(text, destination, out _))
            {
                _propertyCharacterCount = position + length;
                return;
            }

            if (length <= CsvWriter.EscapedShortTextLength && CsvWriter.TryWriteEscapedShortText(text, destination, out int written))
            {
                _propertyCharacterCount = position + written;
                return;
            }

            if (TryStageEncodedText(text, destination, out written))
            {
                _propertyCharacterCount = position + written;
                return;
            }
        }

        if (_propertyFallbackColumn < 0)
        {
            _propertyFallbackColumn = ordinal;
            _propertyFallbackPrefixLength = start - _propertyOrigin;
        }

        StageProjectedText(ordinal, text);
    }

    /// <summary>Escapes text under the default rule and encodes it into the staged row; false when it does not fit.</summary>
    private bool TryStageEncodedText(string text, Span<byte> destination, out int written)
    {
        int capacity = text.Length * 2 + 2;
        char[]? rented = null;
        Span<char> escaped = capacity <= 256 ? stackalloc char[256] : (rented = ArrayPool<char>.Shared.Rent(capacity));
        try
        {
            ReadOnlySpan<char> field = escaped.Slice(0, CsvWriter.WriteEscapedDefault(escaped, text, ','));
            var encoding = _propertyUtf8Writer!.Encoding;
            if (encoding.GetMaxByteCount(field.Length) > destination.Length && encoding.GetByteCount(field) > destination.Length)
            {
                written = 0;
                return false;
            }

            written = encoding.GetBytes(field, destination);
            return true;
        }
        finally
        {
            if (rented != null) ArrayPool<char>.Shared.Return(rented);
        }
    }

    /// <summary>Formats a known invariant scalar as UTF-8 directly into the staged row's available space.</summary>
    private static void StageProjectedScalarUtf8<T>(CsvRowWriter writer, int ordinal, T value)
        where T : ISpanFormattable, IUtf8SpanFormattable
    {
        if (writer._propertyFallbackColumn < 0)
        {
            int start = writer._propertyCharacterCount + (ordinal > 0 ? 1 : 0);
            if (start <= writer._propertyLimit &&
                CsvWriter.TryFormatInvariant(value, writer._propertyBytes.AsSpan(start), out int written))
            {
                if (ordinal > 0) writer._propertyBytes![writer._propertyCharacterCount] = (byte)',';
                writer._propertyCharacterCount = start + written;
                return;
            }

            writer._propertyFallbackColumn = ordinal;
            writer._propertyFallbackPrefixLength = writer._propertyCharacterCount - writer._propertyOrigin;
        }

        StageProjectedScalar(writer, ordinal, value);
    }

    /// <summary>Starts a row in the UTF-8 writer's byte buffer, where it is staged without a copy.</summary>
    private void StartProjectedUtf8Row()
    {
        _propertyBytes = _propertyUtf8Writer!.ReserveUtf8(PropertyStagingLength, out _propertyOrigin);
        _propertyLimit = Math.Min(_propertyBytes.Length, _propertyOrigin + PropertyStagingLength);
        _propertyCharacterCount = _propertyOrigin;
    }

    /// <summary>Commits a staged UTF-8 row with its line terminator.</summary>
    private void WriteProjectedUtf8Row()
    {
        byte[] newLine = _propertyNewLineBytes!;
        int end = _propertyCharacterCount;
        if (newLine.AsSpan().TryCopyTo(_propertyBytes.AsSpan(end)))
        {
            _propertyUtf8Writer!.CommitUtf8(end + newLine.Length - _propertyOrigin);
        }
        else
        {
            _propertyUtf8Writer!.CommitUtf8(end - _propertyOrigin);
            _propertyUtf8Writer.WriteUtf8(newLine);
        }
    }
}
#endif
