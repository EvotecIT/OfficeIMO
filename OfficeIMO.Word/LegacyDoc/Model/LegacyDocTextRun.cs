namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocTextRun {
        internal LegacyDocTextRun(string text, bool bold, bool italic) {
            Text = text;
            Bold = bold;
            Italic = italic;
        }

        internal string Text { get; }

        internal bool Bold { get; }

        internal bool Italic { get; }
    }
}
