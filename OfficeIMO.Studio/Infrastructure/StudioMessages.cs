using System.Globalization;
using System.Text;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Turns engine identifiers and exception text into wording that suits the interface.</summary>
internal static class StudioMessages {
    /// <summary>Removes the ".NET parameter name" suffix that argument exceptions append to their message.</summary>
    internal static string Describe(Exception error) {
        string message = error.Message;
        if (error is ArgumentException { ParamName: { Length: > 0 } parameter })
            message = message.Replace(" (Parameter '" + parameter + "')", string.Empty, StringComparison.Ordinal);
        return message.Trim();
    }

    /// <summary>"render.resource.font-substitution" or "RouteContract" becomes "Font substitution" or "Route contract".</summary>
    internal static string Humanize(string? code) {
        if (string.IsNullOrWhiteSpace(code)) return string.Empty;
        string last = code.Split('.', StringSplitOptions.RemoveEmptyEntries).LastOrDefault() ?? code;
        var words = new List<string>();
        var current = new StringBuilder();
        void Flush() { if (current.Length > 0) { words.Add(current.ToString()); current.Clear(); } }
        for (int index = 0; index < last.Length; index++) {
            char character = last[index];
            if (character is '-' or '_' or ' ') { Flush(); continue; }
            bool startsWord = index > 0 && char.IsUpper(character) &&
                (char.IsLower(last[index - 1]) || index + 1 < last.Length && char.IsLower(last[index + 1]) && char.IsUpper(last[index - 1]));
            if (startsWord) Flush();
            current.Append(character);
        }
        Flush();
        // Acronyms such as PDF stay upper case; other words read as a sentence.
        string text = string.Join(" ", words.Select(word => word.Length > 1 && word.All(char.IsUpper) ? word : word.ToLower(CultureInfo.InvariantCulture)));
        return text.Length == 0 ? code : char.ToUpper(text[0], CultureInfo.CurrentCulture) + text[1..];
    }
}
