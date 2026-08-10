using System.Text;

namespace OfficeIMO.OpenDocument;

/// <summary>
/// Parsed CSS-style font-family syntax used by ODF text properties.
/// </summary>
public sealed class OdfFontFamilySyntax {
    private readonly string[] _families;

    private OdfFontFamilySyntax(string[] families) {
        _families = families;
    }

    /// <summary>Font families in authored fallback order.</summary>
    public IReadOnlyList<string> Families => _families;

    /// <summary>The first authored family.</summary>
    public string PrimaryFamily => _families[0];

    /// <summary>Whether the syntax contains fallback families that a single-name target cannot retain.</summary>
    public bool HasFallbacks => _families.Length > 1;

    /// <summary>Parses an ODF font-family value.</summary>
    /// <exception cref="FormatException">The value is empty or contains malformed list syntax.</exception>
    public static OdfFontFamilySyntax Parse(string value) {
        if (!TryParse(value, out OdfFontFamilySyntax? syntax)) {
            throw new FormatException("The ODF font-family value is not a valid comma-separated family list.");
        }
        return syntax!;
    }

    /// <summary>Tries to parse an ODF font-family value without regular-expression rewriting.</summary>
    public static bool TryParse(string? value, out OdfFontFamilySyntax? syntax) {
        syntax = null;
        if (string.IsNullOrWhiteSpace(value)) return false;

        var families = new List<string>();
        var current = new StringBuilder();
        char quote = '\0';
        bool escaped = false;
        bool closedQuote = false;

        for (int index = 0; index < value!.Length; index++) {
            char character = value[index];
            if (escaped) {
                current.Append(character);
                escaped = false;
                continue;
            }
            if (character == '\\') {
                escaped = true;
                continue;
            }
            if (quote != '\0') {
                if (character == quote) {
                    quote = '\0';
                    closedQuote = true;
                } else {
                    current.Append(character);
                }
                continue;
            }
            if (character == '\'' || character == '"') {
                if (current.ToString().Trim().Length != 0 || closedQuote) return false;
                current.Clear();
                quote = character;
                continue;
            }
            if (character == ',') {
                if (!AddFamily(families, current)) return false;
                closedQuote = false;
                continue;
            }
            if (closedQuote && !char.IsWhiteSpace(character)) return false;
            current.Append(character);
        }

        if (escaped || quote != '\0' || !AddFamily(families, current)) return false;
        syntax = new OdfFontFamilySyntax(families.ToArray());
        return true;
    }

    /// <summary>Formats a stable ODF-compatible family list.</summary>
    public override string ToString() => string.Join(", ", _families.Select(FormatFamily));

    private static bool AddFamily(ICollection<string> families, StringBuilder current) {
        string family = current.ToString().Trim();
        current.Clear();
        if (family.Length == 0) return false;
        families.Add(family);
        return true;
    }

    private static string FormatFamily(string family) {
        if (family.IndexOf(',') < 0 && family.IndexOf('"') < 0 && family.IndexOf('\\') < 0) {
            return family;
        }
        return "\"" + family.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }
}
