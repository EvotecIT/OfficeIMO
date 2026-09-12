using System;

namespace OfficeIMO.Html.Dom;

internal static class HtmlNames {
    // HTML name folding is ASCII-only; Unicode and foreign-content casing are significant.
    internal static string LowerAscii(string name) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        char[]? characters = null;
        for (int index = 0; index < name.Length; index++) {
            char character = name[index];
            if (character < 'A' || character > 'Z') continue;
            characters ??= name.ToCharArray();
            characters[index] = (char)(character + ('a' - 'A'));
        }
        return characters == null ? name : new string(characters);
    }
}
