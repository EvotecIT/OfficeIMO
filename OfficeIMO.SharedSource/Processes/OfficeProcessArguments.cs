using System.Text;

namespace OfficeIMO.Processes.Internal;

/// <summary>Formats argument values for ProcessStartInfo.Arguments on the compatibility targets.</summary>
internal static class OfficeProcessArguments {
    internal static string Join(IEnumerable<string> values) => string.Join(" ", values.Select(QuoteArgument));

    internal static string QuoteArgument(string value) {
        if (value == null) return "\"\"";
        if (value.Length > 0 && value.All(static ch => !char.IsWhiteSpace(ch) && ch != '\"')) return value;
        var builder = new StringBuilder(value.Length + 2);
        builder.Append('\"');
        int slashes = 0;
        foreach (char character in value) {
            if (character == '\\') {
                slashes++;
                continue;
            }
            if (character == '\"') {
                builder.Append('\\', (slashes * 2) + 1);
                builder.Append('\"');
                slashes = 0;
                continue;
            }
            if (slashes > 0) builder.Append('\\', slashes);
            slashes = 0;
            builder.Append(character);
        }
        if (slashes > 0) builder.Append('\\', slashes * 2);
        builder.Append('\"');
        return builder.ToString();
    }
}
