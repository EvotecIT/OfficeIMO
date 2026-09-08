using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.PdfQualityCorpus;

internal sealed class ScanTextAccuracy {
    public int ExpectedCharacters { get; init; }
    public int CharacterEdits { get; init; }
    public int ExpectedWords { get; init; }
    public int WordEdits { get; init; }
    public double CharacterErrorRate => (double)CharacterEdits / Math.Max(1, ExpectedCharacters);
    public double WordErrorRate => (double)WordEdits / Math.Max(1, ExpectedWords);

    internal static ScanTextAccuracy Measure(string expected, string actual) {
        static string Normalize(string text) => Regex.Replace(text.Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();
        expected = Normalize(expected); actual = Normalize(actual);
        Rune[] expectedCharacters = expected.EnumerateRunes().ToArray(), actualCharacters = actual.EnumerateRunes().ToArray();
        string[] expectedWords = expected.Split(' ', StringSplitOptions.RemoveEmptyEntries), actualWords = actual.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return new ScanTextAccuracy { ExpectedCharacters = expectedCharacters.Length, CharacterEdits = Distance(expectedCharacters, actualCharacters),
            ExpectedWords = expectedWords.Length, WordEdits = Distance(expectedWords, actualWords) };
    }

    private static int Distance<T>(T[] expected, T[] actual) where T : IEquatable<T> {
        int[] previous = Enumerable.Range(0, actual.Length + 1).ToArray(), current = new int[actual.Length + 1];
        for (int row = 1; row <= expected.Length; row++) {
            current[0] = row;
            for (int column = 1; column <= actual.Length; column++)
                current[column] = Math.Min(previous[column] + 1, Math.Min(current[column - 1] + 1,
                    previous[column - 1] + (expected[row - 1].Equals(actual[column - 1]) ? 0 : 1)));
            (previous, current) = (current, previous);
        }
        return previous[actual.Length];
    }
}
