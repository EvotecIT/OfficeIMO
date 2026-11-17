using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf.Benchmarks.TextExtraction;

/// <summary>
/// Compares extracted texts pairwise. Texts are NFKC-normalised, control characters are removed and whitespace
/// runs collapse to one space. The similarity ratio is <c>2 * LCS / (|a| + |b|)</c> over the characters with all
/// whitespace removed, so spacing and line-break choices do not count but missing, extra or reordered text does.
/// Very large pairs whose edit distance exceeds the search budget fall back to a word-multiset Dice ratio.
/// </summary>
internal static class TextParity {
    private const long DiffBudget = 400_000_000;

    internal static int Run(string[] args) {
        string output = Path.GetFullPath(Options.Get(args, "--output") ?? throw new ArgumentException("--output is required."));
        string textRoot = Path.Combine(output, "texts");
        string[] pairs = (Options.Get(args, "--pairs") ??
                "officeimo:pdfium,pdfoxide:pdfium,pdfpig:pdfium,officeimo-read:pdfium,officeimo:pdfoxide")
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        // Suites come from every engine named in the pairs, so a run without PDFium, or a suite where one engine
        // failed throughout, still compares the engines that have output.
        string[] suites = pairs.SelectMany(static pair => pair.Split(':'))
            .Distinct(StringComparer.Ordinal)
            .Select(engine => Path.Combine(textRoot, engine))
            .Where(Directory.Exists)
            .SelectMany(Directory.GetDirectories)
            .Select(static directory => Path.GetFileName(directory))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(static suite => suite, StringComparer.Ordinal)
            .ToArray();

        var rows = new List<string>();
        foreach (string suite in suites) {
            var files = new SortedSet<string>(StringComparer.Ordinal);
            foreach (string engine in pairs.SelectMany(static pair => pair.Split(':')).Distinct()) {
                string directory = Path.Combine(textRoot, engine, suite);
                if (Directory.Exists(directory)) {
                    foreach (string file in Directory.EnumerateFiles(directory, "*.txt", SearchOption.AllDirectories)) {
                        files.Add(Path.GetRelativePath(directory, file));
                    }
                }
            }
            var results = new string?[files.Count * pairs.Length];
            string[] fileList = files.ToArray();
            Parallel.For(0, fileList.Length, index => {
                for (int p = 0; p < pairs.Length; p++) {
                    string[] engines = pairs[p].Split(':');
                    string left = Path.Combine(textRoot, engines[0], suite, fileList[index]);
                    string right = Path.Combine(textRoot, engines[1], suite, fileList[index]);
                    if (!File.Exists(left) || !File.Exists(right)) {
                        continue;
                    }
                    string a = File.ReadAllText(left);
                    string b = File.ReadAllText(right);
                    Comparison comparison = Compare(a, b);
                    string file = fileList[index][..^4];
                    results[index * pairs.Length + p] = string.Join(',',
                        CorpusRunner.Csv(suite), CorpusRunner.Csv(file), engines[0], engines[1], comparison.Method,
                        comparison.Similarity.ToString("0.0000", CultureInfo.InvariantCulture),
                        comparison.WordDice.ToString("0.0000", CultureInfo.InvariantCulture),
                        comparison.LeftLength.ToString(CultureInfo.InvariantCulture),
                        comparison.RightLength.ToString(CultureInfo.InvariantCulture),
                        LaxCategory(a, b));
                }
            });
            rows.AddRange(results.Where(static row => row is not null)!);
            Console.WriteLine($"{suite}: compared {fileList.Length} files.");
        }

        string parityPath = Path.Combine(output, "parity.csv");
        using var writer = new StreamWriter(parityPath, append: false, new UTF8Encoding(false));
        writer.WriteLine("suite,file,left,right,method,similarity,word_dice,left_chars,right_chars,lax_category");
        foreach (string row in rows) {
            writer.WriteLine(row);
        }
        Console.WriteLine($"Parity: {parityPath}");
        return 0;
    }

    internal sealed record Comparison(string Method, double Similarity, double WordDice, int LeftLength, int RightLength);

    internal static Comparison Compare(string left, string right) {
        string a = Normalize(left);
        string b = Normalize(right);
        string[] wordsA = a.Length == 0 ? [] : a.Split(' ');
        string[] wordsB = b.Length == 0 ? [] : b.Split(' ');
        double dice = WordDice(wordsA, wordsB);
        string compactA = a.Replace(" ", string.Empty, StringComparison.Ordinal);
        string compactB = b.Replace(" ", string.Empty, StringComparison.Ordinal);
        if (compactA.Length + compactB.Length == 0) {
            return new Comparison("chars", 1, 1, 0, 0);
        }
        int? lcs = LongestCommonSubsequence(compactA, compactB);
        return lcs is int common
            ? new Comparison("chars", 2.0 * common / (compactA.Length + compactB.Length), dice, compactA.Length, compactB.Length)
            : new Comparison("words-bag", dice, dice, compactA.Length, compactB.Length);
    }

    internal static string Normalize(string text) {
        string normalized = PrepareForNormalization(text).Normalize(NormalizationForm.FormKC);
        var builder = new StringBuilder(normalized.Length);
        bool pendingSpace = false;
        foreach (char c in normalized) {
            if (char.IsWhiteSpace(c) || c == '​' || c == '﻿') {
                pendingSpace = builder.Length > 0;
                continue;
            }
            if (char.IsControl(c) || c is '­' or '￾' or '￿') {
                continue;
            }
            if (pendingSpace) {
                builder.Append(' ');
                pendingSpace = false;
            }
            builder.Append(c);
        }
        return builder.ToString();
    }

    /// <summary>
    /// Engines can emit unpaired surrogates and noncharacters for unmappable glyphs, and NFKC rejects both. Unpaired
    /// surrogates become U+FFFD; noncharacters are ignored anyway, so they are dropped before normalisation so the
    /// rest of the text is still normalised.
    /// </summary>
    private static string PrepareForNormalization(string text) {
        StringBuilder? builder = null;
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            bool paired = char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]);
            int codePoint = paired ? char.ConvertToUtf32(c, text[i + 1]) : c;
            int width = paired ? 2 : 1;
            if (IsNoncharacter(codePoint)) {
                builder ??= new StringBuilder(text, 0, i, text.Length);
            } else if (!paired && char.IsSurrogate(c)) {
                builder ??= new StringBuilder(text, 0, i, text.Length);
                builder.Append('�');
            } else {
                builder?.Append(text, i, width);
            }
            i += width - 1;
        }
        return builder?.ToString() ?? text;
    }

    private static bool IsNoncharacter(int codePoint) =>
        codePoint is >= 0xFDD0 and <= 0xFDEF || (codePoint & 0xFFFE) == 0xFFFE;

    private static double WordDice(string[] a, string[] b) {
        if (a.Length + b.Length == 0) {
            return 1;
        }
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string word in a) {
            counts[word] = counts.TryGetValue(word, out int count) ? count + 1 : 1;
        }
        int common = 0;
        foreach (string word in b) {
            if (counts.TryGetValue(word, out int count) && count > 0) {
                counts[word] = count - 1;
                common++;
            }
        }
        return 2.0 * common / (a.Length + b.Length);
    }

    /// <summary>Myers O((N+M)D) LCS length; returns null when the edit distance would exceed the budget.</summary>
    private static int? LongestCommonSubsequence(string a, string b) {
        int n = a.Length, m = b.Length;
        int maxD = (int)Math.Min(n + m, DiffBudget / Math.Max(1, n + m));
        int offset = maxD + 1;
        int[] v = new int[2 * maxD + 3];
        for (int d = 0; d <= maxD; d++) {
            for (int k = -d; k <= d; k += 2) {
                int x = k == -d || (k != d && v[offset + k - 1] < v[offset + k + 1])
                    ? v[offset + k + 1]
                    : v[offset + k - 1] + 1;
                int y = x - k;
                while (x < n && y < m && a[x] == b[y]) {
                    x++;
                    y++;
                }
                v[offset + k] = x;
                if (x >= n && y >= m) {
                    return (n + m - d) / 2;
                }
            }
        }
        return null;
    }

    /// <summary>
    /// The upstream project's own text-equivalence classification, reproduced so its published parity
    /// percentage can be read against stricter metrics. "clean" includes any output whose character set is a
    /// superset of the reference, and any length ratio between 0.8 and 1.2.
    /// </summary>
    internal static string LaxCategory(string candidate, string reference) {
        string ot = StripControl(candidate).Trim();
        string mt = StripControl(reference).Trim();
        if (ot.Length == 0 && mt.Length == 0) {
            return "both_empty";
        }
        if (mt.Length == 0) {
            return "candidate_only";
        }
        if (ot.Length == 0) {
            return "candidate_empty";
        }
        if (ot == mt) {
            return "clean";
        }
        string otCompact = string.Concat(ot.Where(static c => !char.IsWhiteSpace(c)));
        string mtCompact = string.Concat(mt.Where(static c => !char.IsWhiteSpace(c)));
        if (otCompact == mtCompact) {
            return "clean";
        }
        var otChars = new HashSet<char>(otCompact);
        if (mtCompact.Length > 0 && otChars.IsSupersetOf(mtCompact)) {
            return "clean";
        }
        double ratio = (double)ot.Length / mt.Length;
        return ratio switch {
            < 0.5 => "candidate_much_less",
            < 0.8 => "candidate_less",
            > 2.0 => "candidate_much_more",
            > 1.2 => "candidate_more",
            _ => "clean"
        };
    }

    private static string StripControl(string text) =>
        string.Concat(text.Where(static c => c is '\n' or '\r' or '\t' || c >= ' '));
}
