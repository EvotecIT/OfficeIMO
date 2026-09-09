using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Builds a printable onboarding checklist with tasks grouped by the employee journey.</summary>
internal static class OnboardingChecklist {
    internal static void Create(string folder) {
        using WordDocument document = WordDocument.Create(Path.Combine(folder, "example.docx"));
        document.Settings.FontFamily = "Carlito";
        document.Settings.FontSize = 11;
        document.AddParagraph("A good first week").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("Employee onboarding | Team: Customer Operations | Buddy: Morgan Chen");
        document.AddParagraph("Use this checklist to agree who prepares each item and record completion.");

        AddStage("Before arrival", new[] {
            ("Equipment and access", "IT", "Ready"),
            ("Welcome schedule", "Manager", "Ready"),
            ("Buddy introduction", "Team buddy", "Planned")
        });
        AddStage("Day one", new[] {
            ("Workspace and safety tour", "Facilities", "Planned"),
            ("Role and expectations", "Manager", "Planned"),
            ("First service walkthrough", "Team buddy", "Planned")
        });
        AddStage("By the end of week one", new[] {
            ("Complete a supported task", "Employee", "Planned"),
            ("Review questions and blockers", "Manager", "Planned")
        });
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

        void AddStage(string title, (string Task, string Owner, string Status)[] tasks) {
            document.AddParagraph(title).Style = WordParagraphStyles.Heading2;
            WordTable table = document.AddTable(tasks.Length + 1, 3, WordTableStyle.TableGrid);
            string[] headers = { "Task", "Owner", "Status" };
            for (int column = 0; column < 3; column++) {
                table.Rows[0].Cells[column].Paragraphs[0].Text = headers[column];
                table.Rows[0].Cells[column].Paragraphs[0].Bold = true;
                table.Rows[0].Cells[column].ShadingFillColorHex = "E8EEF8";
            }
            for (int row = 0; row < tasks.Length; row++) {
                table.Rows[row + 1].Cells[0].Paragraphs[0].Text = tasks[row].Task;
                table.Rows[row + 1].Cells[1].Paragraphs[0].Text = tasks[row].Owner;
                table.Rows[row + 1].Cells[2].Paragraphs[0].Text = tasks[row].Status;
            }
        }
    }
}
