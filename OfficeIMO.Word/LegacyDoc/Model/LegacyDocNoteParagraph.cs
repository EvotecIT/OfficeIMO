namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocNoteParagraph {
        internal LegacyDocNoteParagraph(IReadOnlyList<LegacyDocTextRun> runs)
            : this(runs, LegacyDocParagraphFormat.Default) {
        }

        internal LegacyDocNoteParagraph(IReadOnlyList<LegacyDocTextRun> runs, LegacyDocParagraphFormat format) {
            Runs = runs.Count == 0
                ? Array.Empty<LegacyDocTextRun>()
                : runs.ToArray();
            Text = string.Concat(Runs.Select(run => run.Text));
            Format = format;
        }

        internal string Text { get; }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }

        internal LegacyDocParagraphFormat Format { get; }
    }
}
