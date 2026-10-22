using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word;

/// <summary>Printed-report labels and invisible identifier wrapping with explicit plain-text normalization.</summary>
public static partial class WordPaperText {
    /// <summary>Separates case and underscore boundaries while retaining spelling and acronyms.</summary>
    /// <param name="value">A property name or an already spaced label.</param>
    /// <returns>A display label with ordinary spaces; never use this method for source identifier values.</returns>
    public static string SplitIdentifier(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        return Regex.Replace(Regex.Replace(value.Replace('_', ' '), "([A-Z]+)([A-Z][a-z])", "$1 $2"),
            "([a-z0-9])([A-Z])", "$1 $2");
    }

    /// <summary>Converts enum or property identifiers to sentence-case labels, preserving acronyms.</summary>
    /// <param name="value">An enum or property identifier.</param>
    /// <returns>A readable label such as "Short term" or "DNS health".</returns>
    public static string HumanizeIdentifier(string value) => SentenceCase(SplitIdentifier(value));

    /// <summary>Converts authored title-case prose to sentence case, preserving uppercase and mixed-case technical terms.</summary>
    /// <remarks>Use only where the caller owns the prose; externally supplied titles may contain proper names whose case must be retained.</remarks>
    /// <param name="value">Title-case prose owned by the report author.</param>
    /// <returns>Sentence-case prose with technical casing retained.</returns>
    public static string SentenceCase(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        string sentence = Regex.Replace(value, @"\b[A-Z][a-z]+\b", match => match.Value.ToLowerInvariant());
        Match firstWord = Regex.Match(sentence, @"\p{L}[\p{L}\p{N}]*");
        if (!firstWord.Success || (firstWord.Value.Any(char.IsUpper) && firstWord.Value.Any(char.IsLower))) return sentence;
        int start = firstWord.Index;
        return sentence.Substring(0, start) + char.ToUpperInvariant(sentence[start]) + sentence.Substring(start + 1);
    }
}
