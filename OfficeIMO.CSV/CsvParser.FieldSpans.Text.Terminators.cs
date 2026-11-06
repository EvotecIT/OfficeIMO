#if NET8_0_OR_GREATER
using System.Buffers;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    private static readonly SearchValues<char> CommaTextFieldTerminators = SearchValues.Create(",\r\n\"");
    private static readonly SearchValues<char> SemicolonTextFieldTerminators = SearchValues.Create(";\r\n\"");
    private static readonly SearchValues<char> TabTextFieldTerminators = SearchValues.Create("\t\r\n\"");

    private static int IndexOfTextFieldTerminator(ReadOnlySpan<char> text, char delimiter)
    {
        switch (delimiter)
        {
            case ',': return text.IndexOfAny(CommaTextFieldTerminators);
            case ';': return text.IndexOfAny(SemicolonTextFieldTerminators);
            case '\t': return text.IndexOfAny(TabTextFieldTerminators);
        }

        var terminator = text.IndexOfAny(delimiter, '\r', '\n');
        var quote = (terminator < 0 ? text : text.Slice(0, terminator)).IndexOf('"');
        return quote >= 0 ? quote : terminator;
    }
}
#endif
