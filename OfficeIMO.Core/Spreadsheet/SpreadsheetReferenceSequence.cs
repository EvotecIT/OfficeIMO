using System;
using System.Collections.Generic;

namespace OfficeIMO.Spreadsheet;

/// <summary>
/// Immutable syntax for an OpenXML-style sequence of spreadsheet references. Separators
/// are recognized only outside quoted worksheet names, so punctuation inside names is
/// never treated as structure.
/// </summary>
public sealed class SpreadsheetReferenceSequence {
    private SpreadsheetReferenceSequence(IReadOnlyList<SpreadsheetRangeReference> references) {
        References = references;
    }

    /// <summary>Gets the parsed references in authored order.</summary>
    public IReadOnlyList<SpreadsheetRangeReference> References { get; }

    /// <summary>Parses a comma- or whitespace-delimited reference sequence.</summary>
    public static SpreadsheetReferenceSequence Parse(string text, SpreadsheetAddressDialect dialect) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (!TryParse(text, dialect, out SpreadsheetReferenceSequence? sequence)) {
            throw new FormatException($"'{text}' is not a valid {dialect} spreadsheet reference sequence.");
        }
        return sequence!;
    }

    /// <summary>Attempts to parse a comma- or whitespace-delimited reference sequence.</summary>
    public static bool TryParse(string? text, SpreadsheetAddressDialect dialect,
        out SpreadsheetReferenceSequence? sequence) {
        sequence = null;
        if (string.IsNullOrWhiteSpace(text)) return false;

        var references = new List<SpreadsheetRangeReference>();
        int cursor = 0;
        while (cursor < text!.Length) {
            while (cursor < text.Length && IsSeparator(text[cursor])) cursor++;
            if (cursor >= text.Length) break;

            int start = cursor;
            bool quoted = false;
            while (cursor < text.Length) {
                char current = text[cursor];
                if (current == '\'') {
                    if (quoted && cursor + 1 < text.Length && text[cursor + 1] == '\'') {
                        cursor += 2;
                        continue;
                    }
                    quoted = !quoted;
                    cursor++;
                    continue;
                }
                if (!quoted && IsSeparator(current)) break;
                cursor++;
            }
            if (quoted) return false;
            string token = text.Substring(start, cursor - start);
            if (!SpreadsheetRangeReference.TryParse(token, dialect, out SpreadsheetRangeReference? reference)) return false;
            references.Add(reference!);
        }

        if (references.Count == 0) return false;
        sequence = new SpreadsheetReferenceSequence(references.AsReadOnly());
        return true;
    }

    private static bool IsSeparator(char value) => value == ',' || char.IsWhiteSpace(value);
}
