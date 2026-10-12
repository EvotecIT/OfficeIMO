#nullable enable

#if NET8_0_OR_GREATER

using System;
using System.Data;

namespace OfficeIMO.Data;

/// <summary>Exposes the current field's text through provider-owned UTF-8 storage.</summary>
/// <remarks>
/// This is an optional capability. A provider may return false for any field that requires decoding,
/// conversion, or materialization. A supporting provider may initialize or update its text cache
/// when a field is requested. The ordinary data-record getters remain available in either case.
/// </remarks>
public interface IDataReaderUtf8TextSource {
    /// <summary>Attempts to borrow the current field's decoded text as UTF-8 bytes.</summary>
    /// <param name="ordinal">The zero-based field ordinal.</param>
    /// <param name="text">The borrowed text when successful; otherwise an empty span.</param>
    /// <returns>True when the field's text can be borrowed from provider-owned storage.</returns>
    /// <remarks>
    /// Consume the span before the next call to Read, ReadAsync, NextResult, NextResultAsync,
    /// Close, or Dispose on its reader. Copy the bytes when they must survive those operations.
    /// An empty text field can return true with an empty span; false does not indicate DBNull.
    /// </remarks>
    /// <exception cref="IndexOutOfRangeException">The ordinal is outside the reader's fields.</exception>
    /// <exception cref="InvalidOperationException">The reader is closed or is not positioned on a row.</exception>
    bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text);
}

/// <summary>Reads borrowed or copied UTF-8 text from ordinary tabular records.</summary>
public static partial class DataReaderUtf8TextExtensions {
    /// <summary>Attempts to borrow text from a record that implements <see cref="IDataReaderUtf8TextSource"/>.</summary>
    /// <param name="record">The record containing the current row.</param>
    /// <param name="ordinal">The zero-based field ordinal.</param>
    /// <param name="text">The borrowed text when successful; otherwise an empty span.</param>
    /// <returns>False when the record or field does not offer borrowed UTF-8 text.</returns>
    /// <remarks>
    /// This method never calls GetValue or converts a value to UTF-8. A supporting provider validates
    /// its cursor state. Consume a successful span before advancing or closing its reader.
    /// </remarks>
    /// <exception cref="ArgumentNullException">The record is null.</exception>
    /// <exception cref="IndexOutOfRangeException">The ordinal is outside the record's fields.</exception>
    public static bool TryGetUtf8Text(this IDataRecord record, int ordinal, out ReadOnlySpan<byte> text) {
        if (record == null) {
            throw new ArgumentNullException(nameof(record));
        }
        if (record is IDataReaderUtf8TextSource source) {
            return source.TryGetUtf8Text(ordinal, out text);
        }
        if ((uint)ordinal >= (uint)record.FieldCount) {
            throw new IndexOutOfRangeException(ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        text = default;
        return false;
    }
}
#endif
