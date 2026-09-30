using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.PdfQualityCorpus;

internal sealed class ScanTextAccuracy {
    public int ExpectedCharacters { get; init; }
    public int CharacterEdits { get; init; }
    public int ExpectedWords { get; init; }
    public int WordEdits { get; init; }
    public int WordDeletions { get; init; }
    public int WordInsertions { get; init; }
    public int WordSubstitutions { get; init; }
    public double CharacterErrorRate => (double)CharacterEdits / Math.Max(1, ExpectedCharacters);
    public double WordErrorRate => (double)WordEdits / Math.Max(1, ExpectedWords);
    public double WordDeletionRate => (double)WordDeletions / Math.Max(1, ExpectedWords);

    internal static string Normalize(string text) => Regex.Replace(text.Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();

    internal static ScanTextAccuracy Measure(string expected, string actual, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        expected = Normalize(expected); actual = Normalize(actual);
        Rune[] expectedCharacters = expected.EnumerateRunes().ToArray(), actualCharacters = actual.EnumerateRunes().ToArray();
        string[] expectedWords = expected.Split(' ', StringSplitOptions.RemoveEmptyEntries), actualWords = actual.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if ((long)expectedCharacters.Length * actualCharacters.Length > 100_000_000)
            throw new InvalidDataException("Text accuracy comparison exceeds the edit-distance work budget.");
        WordAlignment words = AlignWords(expectedWords, actualWords, cancellationToken);
        return new ScanTextAccuracy { ExpectedCharacters = expectedCharacters.Length, CharacterEdits = Distance(expectedCharacters, actualCharacters, cancellationToken),
            ExpectedWords = expectedWords.Length, WordEdits = words.Edits, WordDeletions = words.Deletions,
            WordInsertions = words.Insertions, WordSubstitutions = words.Substitutions };
    }

    private static int Distance<T>(T[] expected, T[] actual, CancellationToken token) where T : IEquatable<T> {
        int[] previous = Enumerable.Range(0, actual.Length + 1).ToArray(), current = new int[actual.Length + 1];
        for (int row = 1; row <= expected.Length; row++) {
            token.ThrowIfCancellationRequested();
            current[0] = row;
            for (int column = 1; column <= actual.Length; column++)
                current[column] = Math.Min(previous[column] + 1, Math.Min(current[column - 1] + 1,
                    previous[column - 1] + (expected[row - 1].Equals(actual[column - 1]) ? 0 : 1)));
            (previous, current) = (current, previous);
        }
        return previous[actual.Length];
    }

    // A deterministic minimum-edit alignment: ties prefer substitution/match, then deletion, then insertion.
    // Deletions include ordering errors; they are not a measure of semantic fact omission.
    private static WordAlignment AlignWords(string[] expected, string[] actual, CancellationToken token) {
        WordAlignment[] previous = new WordAlignment[actual.Length + 1], current = new WordAlignment[actual.Length + 1];
        for (int column = 1; column <= actual.Length; column++) previous[column] = new WordAlignment(0, column, 0);
        for (int row = 1; row <= expected.Length; row++) {
            token.ThrowIfCancellationRequested(); current[0] = new WordAlignment(row, 0, 0);
            for (int column = 1; column <= actual.Length; column++) {
                WordAlignment diagonal = previous[column - 1];
                bool same = expected[row - 1] == actual[column - 1];
                int cost = diagonal.Edits + (same ? 0 : 1);
                if (cost <= previous[column].Edits + 1 && cost <= current[column - 1].Edits + 1)
                    current[column] = new WordAlignment(diagonal.Deletions, diagonal.Insertions, diagonal.Substitutions + (same ? 0 : 1));
                else if (previous[column].Edits <= current[column - 1].Edits)
                    current[column] = new WordAlignment(previous[column].Deletions + 1, previous[column].Insertions, previous[column].Substitutions);
                else current[column] = new WordAlignment(current[column - 1].Deletions, current[column - 1].Insertions + 1, current[column - 1].Substitutions);
            }
            (previous, current) = (current, previous);
        }
        return previous[actual.Length];
    }

    private readonly record struct WordAlignment(int Deletions, int Insertions, int Substitutions) {
        internal int Edits => Deletions + Insertions + Substitutions;
    }
}
