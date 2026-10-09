#nullable enable

#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal sealed partial class CsvUtf8StreamDataReaderRowSource
    {
        internal bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text)
        {
            if (_fallback is not null)
            {
                text = default;
                return false;
            }

            return _visitor.TryGetUtf8Text(ordinal, out text);
        }
    }
}
#endif
