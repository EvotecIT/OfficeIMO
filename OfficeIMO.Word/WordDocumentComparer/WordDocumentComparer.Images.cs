namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void CompareImages(WordDocument source, WordDocument target, WordDocument result) {
            var srcImages = source.GetImages();
            var tgtImages = target.GetImages();
            int count = System.Math.Min(srcImages.Count, tgtImages.Count);

            for (int i = 0; i < count; i++) {
                if (!srcImages[i].SequenceEqual(tgtImages[i])) {
                    WordParagraph p = result.AddParagraph();
                    p.AddDeletedText("[Image]", "Comparer");
                    p.AddInsertedText("[Image]", "Comparer");
                }
            }

            for (int i = count; i < tgtImages.Count; i++) {
                WordParagraph p = result.AddParagraph();
                p.AddInsertedText("[Image]", "Comparer");
            }

            for (int i = count; i < srcImages.Count; i++) {
                WordParagraph p = result.AddParagraph();
                p.AddDeletedText("[Image]", "Comparer");
            }
        }
    }
}
