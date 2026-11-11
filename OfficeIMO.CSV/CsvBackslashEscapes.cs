using System.Text;

namespace OfficeIMO.CSV;

// PostgreSQL COPY text / linear TSV escapes: \t, \n, \r and \\. Other backslash
// sequences (such as PostgreSQL's \N null marker) stay literal on read.
internal static class CsvBackslashEscapes
{
    internal static string Decode(string value)
    {
        int backslash = value.IndexOf('\\');
        if (backslash < 0) return value;

        var builder = new StringBuilder(value.Length);
        builder.Append(value, 0, backslash);
        for (int i = backslash; i < value.Length; i++)
        {
            char current = value[i];
            if (current != '\\' || i + 1 == value.Length)
            {
                builder.Append(current);
                continue;
            }

            char next = value[i + 1];
            char? decoded = next switch
            {
                't' => '\t',
                'n' => '\n',
                'r' => '\r',
                '\\' => '\\',
                _ => null
            };
            if (decoded is { } character)
            {
                builder.Append(character);
                i++;
            }
            else
            {
                builder.Append(current);
            }
        }

        return builder.ToString();
    }

    internal static int IndexOfEncoded(string value) => value.IndexOfAny(EncodedCharacters);

    internal static string Encode(string value)
    {
        int index = IndexOfEncoded(value);
        return index < 0 ? value : Encode(value, index);
    }

    private static string Encode(string value, int firstIndex)
    {
        var builder = new StringBuilder(value.Length + 8);
        builder.Append(value, 0, firstIndex);
        for (int i = firstIndex; i < value.Length; i++)
        {
            switch (value[i])
            {
                case '\\': builder.Append("\\\\"); break;
                case '\t': builder.Append("\\t"); break;
                case '\n': builder.Append("\\n"); break;
                case '\r': builder.Append("\\r"); break;
                default: builder.Append(value[i]); break;
            }
        }

        return builder.ToString();
    }

    private static readonly char[] EncodedCharacters = { '\\', '\t', '\n', '\r' };
}
