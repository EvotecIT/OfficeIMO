using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a short policy brief with reusable heading structure and bullet lists.</summary>
internal static class PolicyBrief {
    internal static void Create(string folder) {
        using WordDocument document = WordDocument.Create(Path.Combine(folder, "example.docx"));
        document.Settings.FontFamily = "Carlito";
        document.Settings.FontSize = 11;
        document.BuiltinDocumentProperties.Title = "Shared workspace guide";
        document.BuiltinDocumentProperties.Subject = "Working agreements";
        document.AddParagraph("Shared workspace guide").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("Version 1.0 | Owner: Workplace team | Review: March 2027");
        document.AddParagraph("A concise working agreement for reserving shared desks and meeting spaces.");

        Section("Who this applies to", "Everyone using the shared office, including visitors and project teams.");
        Section("Reserve the space", "Book the desk or room you need. Include setup time and release unused bookings so another team can use the space.");
        document.AddParagraph("Leave it ready for the next person").Style = WordParagraphStyles.Heading2;
        WordList actions = document.AddList(WordListStyle.Bulleted);
        actions.AddItem("Remove personal items and dispose of waste.");
        actions.AddItem("Return adapters and equipment to the labelled storage area.");
        actions.AddItem("Report damaged equipment to the Workplace team.");

        Section("Exceptions", "Contact the Workplace team for accessible equipment, recurring project space, or a visitor group.");
        Section("Questions and review", "The Workplace team records feedback and reviews the agreement every six months.");
        foreach (WordParagraph paragraph in document.Paragraphs) {
            paragraph.FontFamily = "Carlito";
            paragraph.FontSize = 11;
            paragraph.LineSpacingAfterPoints = 7;
            if (paragraph.Style == WordParagraphStyles.Heading1) {
                paragraph.FontSize = 26;
                paragraph.Bold = true;
                paragraph.LineSpacingAfterPoints = 12;
            } else if (paragraph.Style == WordParagraphStyles.Heading2) {
                paragraph.FontSize = 15;
                paragraph.Bold = true;
                paragraph.LineSpacingBeforePoints = 12;
                paragraph.LineSpacingAfterPoints = 6;
            }
        }
        document.Save();

        void Section(string heading, string body) {
            document.AddParagraph(heading).Style = WordParagraphStyles.Heading2;
            document.AddParagraph(body);
        }
    }
}
