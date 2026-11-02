#nullable enable

#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

public partial class CsvDataReader : IDataReaderUtf8TextSource
{
    /// <summary>
    /// Borrows normalized text only when the current source already holds its exact UTF-8 bytes.
    /// Conversion, missing values and canonical text-parser fallbacks remain unavailable.
    /// </summary>
    public virtual bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text)
    {
        EnsureOpenRow();
        if ((uint)ordinal >= (uint)_columns.Length)
        {
            throw new IndexOutOfRangeException();
        }

        if (_useDirectTextSourceStrings && _utf8StreamTextRowSource is not null)
        {
            return _utf8StreamTextRowSource.TryGetUtf8Text(ordinal, out text);
        }

        text = default;
        return false;
    }
}
#endif
