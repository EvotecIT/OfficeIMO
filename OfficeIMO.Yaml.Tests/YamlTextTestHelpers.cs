using System.Text;

namespace OfficeIMO.Yaml.Tests;

internal static class YamlTextTestHelpers {
    /// <summary>Decodes suite bytes the way the reader does, so offsets index the same text.</summary>
    public static string DecodeSuiteInput(byte[] bytes) {
        using var reader = new StreamReader(new MemoryStream(bytes), new UTF8Encoding(false), detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }

    /// <summary>Formats every event the reader produces, in the suite's notation.</summary>
    public static string Events(YamlReader reader) {
        var lines = new StringBuilder();
        while (reader.Read()) {
            if (reader.TokenType == YamlTokenType.Comment) {
                lines.Append("#").Append(reader.GetString()).Append(' ').Append(reader.Start).Append('\n');
                continue;
            }
            lines.Append(Conformance.OfficeIMOYamlConformanceEngine.Convert(reader)).Append(' ')
                .Append(reader.Start.Offset).Append('-').Append(reader.End.Offset).Append('\n');
        }

        return lines.ToString();
    }
}

/// <summary>A text reader that returns at most a few characters per call, forcing every token across buffer refills.</summary>
internal sealed class TrickleReader : TextReader {
    private readonly string _text;
    private readonly int _chunk;
    private int _position;

    public TrickleReader(string text, int chunk) {
        _text = text;
        _chunk = chunk;
    }

    public override int Read(char[] buffer, int index, int count) {
        int n = Math.Min(Math.Min(count, _chunk), _text.Length - _position);
        _text.CopyTo(_position, buffer, index, n);
        _position += n;
        return n;
    }
}
