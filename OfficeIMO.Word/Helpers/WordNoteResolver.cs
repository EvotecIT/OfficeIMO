using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Resolves referenced ordinary note bodies independently of their numeric identifiers.</summary>
    internal static class WordNoteResolver {
        /// <summary>Distinguishes ordinary notes, whose type defaults to normal, from special separator stories.</summary>
        internal static bool IsOrdinary(FootnoteEndnoteType note) {
            return note.Type == null || note.Type.Value == FootnoteEndnoteValues.Normal;
        }

        /// <summary>Finds an ordinary footnote for a present identifier; zero is a valid identifier.</summary>
        internal static Footnote? FindFootnote(MainDocumentPart? mainPart, long? referenceId) {
            if (!referenceId.HasValue) return null;
            return mainPart?.FootnotesPart?.Footnotes?.Elements<Footnote>()
                .FirstOrDefault(note => note.Id?.Value == referenceId.Value && IsOrdinary(note));
        }

        /// <summary>Finds an ordinary endnote for a present identifier; zero is a valid identifier.</summary>
        internal static Endnote? FindEndnote(MainDocumentPart? mainPart, long? referenceId) {
            if (!referenceId.HasValue) return null;
            return mainPart?.EndnotesPart?.Endnotes?.Elements<Endnote>()
                .FirstOrDefault(note => note.Id?.Value == referenceId.Value && IsOrdinary(note));
        }
    }
}
