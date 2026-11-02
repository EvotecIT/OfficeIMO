using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates an editable statement of work with a decision, supporting table and named next actions.</summary>
    internal static class StatementOfWork {
        internal static void Create(string folder) {
            using var document = WordDocument.Create(Path.Combine(folder, "example.docx"));
            document.Settings.FontSize = global::OfficeIMO.OfficeLength.Points(11);
            document.BuiltinDocumentProperties.Title = "Statement of work";
            document.AddParagraph("Statement of work").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("NORTHWIND / DOCUMENT AUTOMATION / SOW-018");
            document.AddParagraph("Deliver an editable reporting workflow with a clear boundary between implementation and acceptance. This example is a scope record, not a legal template.");
            document.AddParagraph("Deliverables and acceptance").Style = WordParagraphStyles.Heading2;
            string[][] rows = {
                new[] { "Deliverable", "Owner", "Acceptance evidence", "Target" },
                new[] { "Data mapping", "Analyst", "Five agreed sample records", "Week 1" },
                new[] { "Report generator", "Engineer", "DOCX and PDF from each record", "Week 3" },
                new[] { "Operator handover", "Delivery lead", "Owner completes an unaided run", "Week 4" }
            };
            var table = document.AddTable(rows.Length, 4, WordTableStyle.TableGrid);
            table.SetWidthPercentage(100);
            table.SetColumnWidthsPercentage(25, 20, 30, 25);
            for (int row = 0; row < rows.Length; row++) {
                for (int column = 0; column < 4; column++) {
                    var cell = table.Rows[row].Cells[column];
                    cell.MarginTop = global::OfficeIMO.OfficeLength.Centimeters(0.14D);
                    cell.MarginBottom = global::OfficeIMO.OfficeLength.Centimeters(0.14D);
                    cell.MarginLeft = global::OfficeIMO.OfficeLength.Centimeters(0.16D);
                    cell.MarginRight = global::OfficeIMO.OfficeLength.Centimeters(0.16D);
                    cell.ShadingFillColorHex = row == 0 ? "17365D" : row % 2 == 0 ? "F1F5F9" : "FFFFFF";
                    var paragraph = cell.Paragraphs[0];
                    paragraph.Text = rows[row][column];
                    paragraph.Bold = row == 0;
                    paragraph.ColorHex = row == 0 ? "FFFFFF" : "243247";
                    paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(10);
                    paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(2D);
                }
            }
            document.AddParagraph("Scope and change control").Style = WordParagraphStyles.Heading2;
            document.AddParagraph("Includes one approved data source, one report layout and a repeatable export command. Hosting, new data connectors and production rollout are excluded. Record changes with their effect on cost, dates and acceptance before starting them.");
            document.AddParagraph("Next actions").Style = WordParagraphStyles.Heading2;
            var actions = document.AddList(WordListStyle.Bulleted);
            actions.AddItem("Customer owner: approve the sample data and layout.");
            actions.AddItem("Delivery lead: record acceptance for each deliverable.");
            actions.AddItem("Both owners: agree any scope change before implementation.");
            foreach (var paragraph in document.Paragraphs) {
                bool title = paragraph.Style == WordParagraphStyles.Heading1;
                bool heading = paragraph.Style == WordParagraphStyles.Heading2;
                paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(title ? 26 : heading ? 15 : 11);
                paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(heading ? 14 : 0);
                paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(title ? 14 : heading ? 8 : 7);
                if (title || heading) {
                    paragraph.ColorHex = "17365D";
                    paragraph.Bold = true;
                    paragraph.KeepWithNext = true;
                }
            }
            WorkflowFooter.Add(document, "Statement of work");
            document.Save();
        }
    }
}
