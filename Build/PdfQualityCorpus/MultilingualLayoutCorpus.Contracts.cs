namespace OfficeIMO.PdfQualityCorpus;

internal static partial class MultilingualLayoutCorpus {
    internal static void VerifyScoringContract() {
        ScanTextAccuracy deletion = ScanTextAccuracy.Measure("red green blue", "red blue");
        ScanTextAccuracy insertion = ScanTextAccuracy.Measure("red blue", "red green blue");
        ScanTextAccuracy substitution = ScanTextAccuracy.Measure("red green blue", "red pink blue");
        if (deletion.WordEdits != 1 || deletion.WordDeletions != 1 || insertion.WordInsertions != 1 ||
            substitution.WordSubstitutions != 1 || ScanTextAccuracy.Measure("e\u0301", "é").CharacterEdits != 0)
            throw new InvalidOperationException("OCR scoring lost independent edit operations or Unicode normalization.");
        using (var canceled = new CancellationTokenSource()) {
            canceled.Cancel();
            try { ScanTextAccuracy.Measure("A", "A", canceled.Token); throw new InvalidOperationException("Canceled OCR scoring completed."); }
            catch (OperationCanceledException) { }
        }
        try {
            ScanTextAccuracy.Measure(new string('a', 10_001), new string('b', 10_001));
            throw new InvalidOperationException("OCR scoring exceeded its work budget.");
        } catch (InvalidDataException) { }
        string[][] expected = { new[] { "Name", "Count" }, new[] { "Alfa", "24" }, new[] { "Beta", "12" } };
        Verify(expected.Length, new[] { expected }, "exact ordered table");
        Verify(1, new[] { new[] { expected[0], expected[2], expected[1] } }, "reordered rows");
        Verify(0, new[] { new[] { expected[0], expected[1] }, new[] { expected[2] } }, "rows split between tables");
        Verify(0, new[] { expected.Concat(new[] { expected[1] }).ToArray() }, "extra duplicate row");
        Verify(2, new[] { new[] { expected[0], expected[1], expected[1] } }, "duplicate replacing a row");
        Verify(0, new[] { expected, expected }, "duplicate tables");
        Verify(0, Array.Empty<string[][]>(), "missing table");
        Verify(3, new[] { new[] { new[] { " Name ", "Count" }, expected[1], expected[2] } }, "whitespace normalization");
        var strict = new LayoutAcceptance(0, 0, true, true);
        ScanTextAccuracy exact = ScanTextAccuracy.Measure("A B", "A B");
        if (!strict.IsSatisfied(exact, 2, 2, 1, 1, 0, 0, 0) ||
            strict.IsSatisfied(exact, 2, 2, 0, 1, 0, 0, 0) ||
            strict.IsSatisfied(exact, 1, 2, 1, 1, 0, 0, 0) ||
            strict.IsSatisfied(exact, 2, 2, 1, 1, 0, 0, 1) ||
            strict.IsSatisfied(exact, 2, 2, 1, 1, 2, 3, 1) ||
            strict.IsSatisfied(ScanTextAccuracy.Measure("A B", "A C"), 2, 2, 1, 1, 0, 0, 0))
            throw new InvalidOperationException("Layout qualification accepted incomplete order, incorrect text, or unexpected tables.");

        void Verify(int score, string[][][] actual, string contract) {
            int measured = CountExactTableRows(expected, actual);
            if (measured != score)
                throw new InvalidOperationException($"Layout scoring failed for {contract}: expected {score}, got {measured}.");
        }
    }
}
