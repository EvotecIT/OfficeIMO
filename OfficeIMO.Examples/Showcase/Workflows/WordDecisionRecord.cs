using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Captures a decision, the alternatives considered, and the conditions for revisiting it.</summary>
internal static class WordDecisionRecord {
    internal static void Create(string folder) {
        using WordDocument document = WordDocument.Create(Path.Combine(folder, "example.docx"));
        document.Settings.FontFamily = "Carlito";
        document.Settings.FontSize = 11;
        document.AddParagraph("Decision record / DR-014").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("Standardize the weekly service report");
        document.AddParagraph("Status: Accepted | Decision owner: Operations | Date: 14 September 2026");
        document.AddParagraph("Context").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("Teams currently prepare separate reports. Reviewers need consistent measures and a clear place to find the supporting data.");

        document.AddParagraph("Options considered").Style = WordParagraphStyles.Heading2;
        string[,] options = {
            { "Option", "Strength", "Trade-off" },
            { "Free-form documents", "Flexible for each team", "Hard to compare" },
            { "Shared document template", "Consistent review structure", "Requires an owner" },
            { "Dashboard only", "Always available", "Harder to annotate offline" }
        };
        WordTable table = document.AddTable(4, 3, WordTableStyle.TableGrid);
        for (int row = 0; row < 4; row++) {
            for (int column = 0; column < 3; column++) {
                table.Rows[row].Cells[column].Paragraphs[0].Text = options[row, column];
                table.Rows[row].Cells[column].Paragraphs[0].Bold = row == 0;
                if (row == 0) table.Rows[row].Cells[column].ShadingFillColorHex = "E8EEF8";
            }
        }
        document.AddParagraph("Decision").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("Use a shared document template, with links to the source dashboard for detailed investigation.");
        document.AddParagraph("Revisit when").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("The review audience changes, the reporting measures change, or the preparation cost exceeds one hour per team.");
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
            styledTable.SetColumnWidthsPercentage(30, 30, 40);
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
