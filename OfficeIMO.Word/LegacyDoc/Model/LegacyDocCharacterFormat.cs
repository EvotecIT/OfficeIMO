namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocCharacterFormat {
        internal LegacyDocCharacterFormat(bool bold, bool italic) {
            Bold = bold;
            Italic = italic;
        }

        internal bool Bold { get; }

        internal bool Italic { get; }

        internal static LegacyDocCharacterFormat Default { get; } = new LegacyDocCharacterFormat(false, false);
    }
}
