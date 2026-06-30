namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocNoteParagraph {
        internal LegacyDocNoteParagraph(IReadOnlyList<LegacyDocTextRun> runs) {
            Runs = runs.Count == 0
                ? Array.Empty<LegacyDocTextRun>()
                : runs.ToArray();
            Text = string.Concat(Runs.Select(run => run.Text));
        }

        internal string Text { get; }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }
    }
}
