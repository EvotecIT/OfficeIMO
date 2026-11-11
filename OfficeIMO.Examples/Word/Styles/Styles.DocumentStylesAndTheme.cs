using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

internal static class Styles {
    /// <summary>
    /// Builds a corporate template with its own theme and styles, then creates a report that imports both and
    /// uses them for paragraphs, runs and a table.
    /// </summary>
    internal static void Example_DocumentStylesAndTheme(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating a template with its own styles and theme, then a report that uses them");
        string templatePath = Path.Combine(folderPath, "CorporateStyles.dotx");
        string reportPath = Path.Combine(folderPath, "StyledReport.docx");

        using (WordDocument template = WordDocument.Create(templatePath)) {
            // The theme drives every style below that selects a theme colour or theme font.
            template.Theme.Name = "Contoso";
            template.Theme.SetColor(WordThemeColor.Accent1, OfficeColor.Parse("#0B5394"));
            template.Theme.SetColor(WordThemeColor.Accent2, OfficeColor.Parse("#E69138"));
            template.Theme.MajorFont = "Georgia";
            template.Theme.MinorFont = "Verdana";

            WordStyle normal = template.Styles.GetOrAdd(WordParagraphStyles.Normal);
            normal.ThemeFont = WordThemeFont.Minor;
            normal.LineSpacingAfter = OfficeLength.Points(6);

            WordStyle heading = template.Styles.GetOrAdd(WordParagraphStyles.Heading1);
            heading.ThemeFont = WordThemeFont.Major;
            heading.ThemeColor = WordThemeColor.Accent1;
            heading.Borders.BottomStyle = WordBorderStyle.Single;
            heading.Borders.BottomSize = OfficeLength.Points(1);
            heading.Borders.BottomThemeColor = WordThemeColor.Accent1;

            WordStyle note = template.Styles.AddParagraphStyle("CalloutNote", "Callout Note");
            note.IndentationBefore = OfficeLength.Centimeters(1);
            note.ShadingFillColor = OfficeColor.Parse("#FFF2CC");
            note.Italic = true;

            WordStyle noteBody = template.Styles.AddParagraphStyle("CalloutNoteBody", "Callout Note Body");
            noteBody.BasedOnStyleId = note.StyleId;
            noteBody.Italic = false;
            note.NextStyleId = noteBody.StyleId;

            WordStyle keyTerm = template.Styles.AddCharacterStyle("KeyTerm", "Key Term");
            keyTerm.Bold = true;
            keyTerm.ThemeColor = WordThemeColor.Accent2;

            WordStyle grid = template.Styles.AddTableStyle("ContosoGrid", "Contoso Grid");
            grid.SetTableBorders(WordBorderStyle.Single, OfficeLength.Points(0.5), OfficeColor.Parse("#9FC5E8"));
            WordTableStyleRegionFormat header = grid.GetTableRegion(WordTableStyleRegion.FirstRow);
            header.Bold = true;
            header.Color = OfficeColor.Parse("#FFFFFF");
            header.FillColor = OfficeColor.Parse("#0B5394");
            grid.GetTableRegion(WordTableStyleRegion.BandedRows).FillColor = OfficeColor.Parse("#EAF1FB");

            template.Save();
        }

        using (WordDocument template = WordDocument.Load(templatePath, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly }))
        using (WordDocument report = WordDocument.Create(reportPath)) {
            report.Theme.Import(template);
            WordStyleImportReport imported = report.Styles.Import(template, WordStyleConflictMode.Overwrite);
            Console.WriteLine($"    Styles added: {imported.Added.Count}, replaced: {imported.Replaced.Count}");

            report.AddParagraph("Quarterly review").SetStyle(report.Styles["Heading1"]);
            WordParagraph intro = report.AddParagraph("Revenue grew because of ");
            intro.AddText("recurring contracts").SetCharacterStyle(report.Styles["KeyTerm"]);
            intro.AddText(" in every region.");
            report.AddParagraph("Figures are unaudited.").SetStyle(report.Styles["CalloutNote"]);

            WordTable table = report.AddTable(3, 2);
            table.SetStyle(report.Styles["ContosoGrid"]);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Revenue";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "North";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "1.2M";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "South";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "0.9M";

            foreach (string styleId in imported.Added) {
                WordStyle style = report.Styles[styleId];
                Console.WriteLine($"      added {style.Type,-9} {style.StyleId} ({style.Name})");
            }

            report.Save();
            if (openWord) report.OpenInApplication();
        }
    }
}
