namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>Maps Word cell quarter turns to PDF's upward-positive coordinate system.</summary>
        private static int GetNativeCellTextRotation(WordTextDirection? direction) => direction switch {
            WordTextDirection.TopToBottomRightToLeft or WordTextDirection.TopToBottomRightToLeft2010 or
            WordTextDirection.TopToBottomRightToLeftRotated or WordTextDirection.TopToBottomRightToLeftRotated2010 or
            WordTextDirection.TopToBottomLeftToRightRotated or WordTextDirection.TopToBottomLeftToRightRotated2010 => -90,
            WordTextDirection.BottomToTopLeftToRight or WordTextDirection.BottomToTopLeftToRight2010 => 90,
            _ => 0
        };
    }
}
