using OfficeIMO.Word;

internal static class MergeSplitSmoke {
    internal static void Run() {
        using WordDocument destination = WordDocument.Create();
        destination.AddParagraph("Cover").Style = WordParagraphStyles.Heading1;
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Appendix").Style = WordParagraphStyles.Heading1;
        source.AddParagraph("Appendix text").AddFootNote("Appendix note");

        destination.AppendDocument(source);
        IReadOnlyList<WordDocument> sections = destination.SplitBySection();
        IReadOnlyList<WordDocument> chapters = destination.SplitByHeading();
        try {
            if (sections.Count != 2 || chapters.Count != 2 || chapters[1].FootNotes.Count != 1 || chapters[0].FootNotes.Count != 0) {
                throw new InvalidOperationException("Appended content did not split back into its sections and chapters.");
            }
        } finally {
            foreach (WordDocument piece in sections.Concat(chapters)) piece.Dispose();
        }
        Console.WriteLine("PASS | Append a document and split it by section and heading");
    }
}
