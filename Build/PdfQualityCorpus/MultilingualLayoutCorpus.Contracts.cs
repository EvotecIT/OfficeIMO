namespace OfficeIMO.PdfQualityCorpus;

internal static partial class MultilingualLayoutCorpus {
    internal static void VerifyScoringContract() {
        string[][] expected = { new[] { "Name", "Count" }, new[] { "Alfa", "24" }, new[] { "Beta", "12" } };
        Verify(expected.Length, new[] { expected }, "exact ordered table");
        Verify(1, new[] { new[] { expected[0], expected[2], expected[1] } }, "reordered rows");
        Verify(0, new[] { new[] { expected[0], expected[1] }, new[] { expected[2] } }, "rows split between tables");
        Verify(0, new[] { expected.Concat(new[] { expected[1] }).ToArray() }, "extra duplicate row");
        Verify(2, new[] { new[] { expected[0], expected[1], expected[1] } }, "duplicate replacing a row");
        Verify(0, new[] { expected, expected }, "duplicate tables");
        Verify(0, Array.Empty<string[][]>(), "missing table");
        Verify(3, new[] { new[] { new[] { " Name ", "Count" }, expected[1], expected[2] } }, "whitespace normalization");

        void Verify(int score, string[][][] actual, string contract) {
            int measured = CountExactTableRows(expected, actual);
            if (measured != score)
                throw new InvalidOperationException($"Layout scoring failed for {contract}: expected {score}, got {measured}.");
        }
    }
}
