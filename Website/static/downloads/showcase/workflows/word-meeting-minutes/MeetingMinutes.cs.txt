using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Turns meeting decisions and actions into an editable record with a running header.</summary>
internal static class MeetingMinutes {
    internal static void Create(string folder) {
        using WordDocument document = WordDocument.Create(Path.Combine(folder, "example.docx"));
        document.Settings.FontFamily = "Carlito";
        document.Settings.FontSize = 11;
        document.AddHeadersAndFooters();
        document.Header!.Default!.AddParagraph("NORTHWIND / DELIVERY REVIEW");
        document.Footer!.Default!.AddParagraph("Sample meeting record | Internal working copy");
        document.AddParagraph("Delivery review minutes").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("14 September 2026 | Facilitator: Jordan Lee | Duration: 30 minutes");
        document.AddParagraph("Attendees: Product, Engineering, Operations and Support.");

        document.AddParagraph("Decisions").Style = WordParagraphStyles.Heading2;
        WordList decisions = document.AddList(WordListStyle.Bulleted);
        decisions.AddItem("Start the pilot with the support team before inviting all departments.");
        decisions.AddItem("Keep the current escalation route during the first two weeks.");
        document.AddParagraph("Actions to follow up").Style = WordParagraphStyles.Heading2;
        string[,] values = {
            { "Action", "Owner", "Due" },
            { "Confirm pilot participants", "Product", "18 September" },
            { "Publish the support guide", "Support", "21 September" },
            { "Exercise recovery", "Operations", "23 September" }
        };
        WordTable actions = document.AddTable(4, 3, WordTableStyle.TableGrid);
        for (int row = 0; row < 4; row++) {
            for (int column = 0; column < 3; column++) {
                actions.Rows[row].Cells[column].Paragraphs[0].Text = values[row, column];
                actions.Rows[row].Cells[column].Paragraphs[0].Bold = row == 0;
                if (row == 0) actions.Rows[row].Cells[column].ShadingFillColorHex = "E8EEF8";
            }
        }
        document.AddParagraph("Next meeting").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("28 September: review pilot feedback and decide whether to expand access.");
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
        foreach (WordTable styledTable in document.Tables) {
            styledTable.SetWidthPercentage(100);
            styledTable.SetColumnWidthsPercentage(50, 25, 25);
            foreach (var row in styledTable.Rows) {
                foreach (var cell in row.Cells) {
                    cell.MarginTopCentimeters = 0.10;
                    cell.MarginBottomCentimeters = 0.10;
                    foreach (var paragraph in cell.Paragraphs) {
                        paragraph.FontFamily = "Carlito";
                        paragraph.FontSize = 11;
                        paragraph.LineSpacingAfterPoints = 3;
                    }
                }
            }
        }
        document.Save();
    }
}
