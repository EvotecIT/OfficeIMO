using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Builds a research brief with editable footnotes that retain the evidence behind each claim.</summary>
internal static class ResearchBrief {
    internal static void Create(string folder) {
        using WordDocument document = WordDocument.Create(Path.Combine(folder, "example.docx"));
        document.Settings.FontFamily = "Carlito";
        document.Settings.FontSize = 11;
        document.BuiltinDocumentProperties.Title = "Customer research brief";
        document.AddParagraph("Customer research brief").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("NORTHWIND / REQUEST PORTAL / SYNTHETIC RESEARCH DATA");
        document.AddParagraph("Question: where do people lose confidence after submitting a request?");
        document.AddParagraph("What we observed").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("Seven of ten participants looked for an acknowledgement before leaving the portal.")
            .AddFootNote("Illustrative moderated study, ten participants, task 2. Counts describe this sample only.");
        document.AddParagraph("Four participants reopened their request because the next step was unclear.")
            .AddFootNote("Illustrative observation log: sessions 02, 04, 07 and 09. Reopening is not a measure of task failure.");
        document.AddParagraph("Recommendation").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("Show the request identifier, owning team and next expected update immediately after submission.");
        WordList actions = document.AddList(WordListStyle.Bulleted);
        actions.AddItem("Prototype a confirmation panel with a visible request identifier.");
        actions.AddItem("Test whether participants can explain what happens next.");
        actions.AddItem("Measure repeat contact separately from page revisits.");
        document.AddParagraph("Limits and next study").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("These are invented observations for the example. A small moderated sample can reveal usability problems; it does not establish population-wide rates.");
        foreach (var paragraph in document.Paragraphs) {
            paragraph.FontFamily = "Carlito";
            paragraph.FontSize = paragraph.Style == WordParagraphStyles.Heading1 ? 26 : paragraph.Style == WordParagraphStyles.Heading2 ? 15 : 11;
            paragraph.LineSpacingAfterPoints = 9;
        }
        document.Save();
    }
}
