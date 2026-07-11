#nullable enable

using System;

namespace OfficeIMO.CSV;

internal interface ICsvDataReaderTextRowSource : IDisposable
{
    bool Read();

#if NET8_0_OR_GREATER
    ReadOnlySpan<char> GetSpan(int ordinal);
#endif

    string GetString(int ordinal);

    bool IsNull(int ordinal, string? nullValue);

    int CopyStringValues(object[] values, int count, string? nullValue);
}
