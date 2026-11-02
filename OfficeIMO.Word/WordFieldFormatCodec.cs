namespace OfficeIMO.Word;

/// <summary>Separates public enum names from the case-sensitive switches persisted in Word field instructions.</summary>
internal static class WordFieldFormatCodec {
    internal static string ToInstruction(WordFieldFormat value) => value switch {
        WordFieldFormat.UpperRoman => "Roman",
        WordFieldFormat.LowerRoman => "roman",
        WordFieldFormat.UpperLetter => "ALPHABETICAL",
        WordFieldFormat.LowerLetter => "Alphabetical",
        WordFieldFormat.MergeFormat => "Mergeformat",
        _ => value.ToString()
    };

    internal static bool TryParse(string instruction, out WordFieldFormat value) {
        switch (instruction) {
            case "Roman": value = WordFieldFormat.UpperRoman; return true;
            case "roman": value = WordFieldFormat.LowerRoman; return true;
            case "ALPHABETICAL": value = WordFieldFormat.UpperLetter; return true;
            case "Alphabetical": value = WordFieldFormat.LowerLetter; return true;
        }
        // Preserve the original parser's case-insensitive fallback without confusing the two case-sensitive forms.
        if (string.Equals(instruction, "Roman", StringComparison.OrdinalIgnoreCase)) { value = WordFieldFormat.UpperRoman; return true; }
        if (string.Equals(instruction, "Alphabetical", StringComparison.OrdinalIgnoreCase)) { value = WordFieldFormat.LowerLetter; return true; }
        bool parsed = Enum.TryParse(instruction, false, out value) || Enum.TryParse(instruction, true, out value);
        return parsed && value is not WordFieldFormat.UpperRoman and not WordFieldFormat.LowerRoman
            and not WordFieldFormat.UpperLetter and not WordFieldFormat.LowerLetter;
    }
}
