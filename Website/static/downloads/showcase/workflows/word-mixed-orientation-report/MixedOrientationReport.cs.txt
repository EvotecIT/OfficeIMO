using OfficeIMO.Drawing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Keeps the narrative portrait and places a wide decision matrix in its own landscape section.</summary>
internal static class MixedOrientationReport {
    internal static void Create(string folder) {
        using WordDocument document = WordDocument.Create(Path.Combine(folder, "example.docx"));
        document.Settings.FontFamily = "Carlito";
        document.Settings.FontSize = 11;
        document.AddParagraph("Supplier evaluation").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("NORTHWIND / PROCUREMENT / ILLUSTRATIVE OPTIONS");
        document.AddParagraph("Decision to make").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("Choose a support platform for a 25-person operations team. The selected service must support clear ownership, exportable records and a tested recovery process.");
        document.AddParagraph("Evaluation approach").Style = WordParagraphStyles.Heading2;
        var list = document.AddList(WordListStyle.Bulleted);
        list.AddItem("Demonstrate the same request journey with every supplier.");
        list.AddItem("Record evidence separately from claims made during the demonstration.");
        list.AddItem("Keep unresolved questions visible before signing.");
        document.AddParagraph("Recommendation").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("Take Cedar to a controlled pilot. Confirm export completeness and escalation coverage before making a final selection.");
        var section = document.AddSection();
        section.PageOrientation = OfficePageOrientation.Landscape;
        document.AddParagraph("Evidence comparison").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("A landscape section gives each criterion enough room for a useful explanation.");
        string[,] rows = {
            { "Option", "Request ownership", "Data export", "Recovery evidence", "Commercial question" },
            { "Cedar", "Named owner and queue history shown", "CSV and attachment export demonstrated", "Restore exercise supplied", "Confirm support response window" },
            { "Birch", "Queue ownership shown; individual handoff unclear", "CSV shown; attachments untested", "Procedure supplied; exercise pending", "Confirm annual uplift" },
            { "Maple", "Ownership and escalation demonstrated", "Full export pending pilot", "Restore exercise supplied", "Confirm exit assistance" }
        };
        var table = document.AddTable(4, 5, WordTableStyle.TableGrid);
        table.SetWidthPercentage(100);
        table.SetColumnWidthsPercentage(12, 24, 22, 22, 20);
        for (int row = 0; row < 4; row++) for (int column = 0; column < 5; column++) {
            var cell = table.Rows[row].Cells[column];
            cell.Paragraphs[0].Text = rows[row, column];
            cell.Paragraphs[0].Bold = row == 0;
            cell.MarginTopCentimeters = cell.MarginBottomCentimeters = 0.15;
            if (row == 0) cell.ShadingFillColorHex = "E8EEF8";
        }
        foreach (var paragraph in document.Paragraphs) {
            paragraph.FontFamily = "Carlito";
            paragraph.FontSize = paragraph.Style == WordParagraphStyles.Heading1 ? 26 : paragraph.Style == WordParagraphStyles.Heading2 ? 15 : 11;
            paragraph.LineSpacingAfterPoints = 9;
        }
        document.Save();
    }
}
