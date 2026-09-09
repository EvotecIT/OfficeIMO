using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates an editable project charter with scope, milestones, and named owners.</summary>
internal static class ProjectCharter {
    internal static void Create(string folder) {
        using WordDocument document = WordDocument.Create(Path.Combine(folder, "example.docx"));
        document.Settings.FontFamily = "Carlito";
        document.Settings.FontSize = 11;
        document.BuiltinDocumentProperties.Title = "Project charter";
        document.AddParagraph("Project charter").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("NORTHWIND / CUSTOMER PORTAL / VERSION 1.0");
        document.AddParagraph("Give the delivery team one agreed reference for the outcome, scope, and acceptance decision.");

        document.AddParagraph("Outcome and boundaries").AddBookmark("Scope").Style = WordParagraphStyles.Heading2;
        WordList scope = document.AddList(WordListStyle.Bulleted);
        scope.AddItem("Provide a single place to submit and track service requests.");
        scope.AddItem("Include request history and a clear escalation path.");
        scope.AddItem("Keep billing and account administration outside this release.");

        document.AddParagraph("Delivery milestones").Style = WordParagraphStyles.Heading2;
        string[,] values = {
            { "Milestone", "Owner", "Acceptance evidence" },
            { "Discovery", "Product", "Reviewed service journey" },
            { "Pilot", "Engineering", "Ten completed pilot requests" },
            { "Launch", "Operations", "Support and recovery walkthrough" }
        };
        WordTable table = document.AddTable(4, 3, WordTableStyle.TableGrid);
        for (int row = 0; row < 4; row++) {
            for (int column = 0; column < 3; column++) {
                var paragraph = table.Rows[row].Cells[column].Paragraphs[0];
                paragraph.Text = values[row, column];
                paragraph.Bold = row == 0;
                if (row == 0) table.Rows[row].Cells[column].ShadingFillColorHex = "E8EEF8";
            }
        }
        document.AddParagraph("Success measure").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("At least 90% of pilot requests reach the right team without manual reassignment.");
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
            styledTable.SetColumnWidthsPercentage(25, 22, 53);
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
