namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static void RequireOnixLanguageCode(string value, string name) {
        RequireOnixText(value, name);
        if (value.Length != 3 || value.Any(c => c < 'a' || c > 'z'))
            throw new ArgumentException("Supply a three-letter ONIX list 74 language code.", name);
    }

    // Repeated parallel text needs an explicit language on every instance, within its own composite.
    internal static void RequireOnixTranslationLanguage(string? language, int count, string name) {
        if (count > 1 && language == null)
            throw new ArgumentException("Repeated ONIX translations require an explicit language on every entry.", name);
        if (language != null) RequireOnixLanguageCode(language, name);
    }
}
