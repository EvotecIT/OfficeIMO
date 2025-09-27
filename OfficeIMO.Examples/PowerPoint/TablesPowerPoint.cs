using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates table cell manipulation and row/column management.
    /// </summary>
    public static class TablesPowerPoint {
        public static void Example_PowerPointTables(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Table operations");
            string filePath = Path.Combine(folderPath, "Table Operations.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            const long width = 5_000_000L;
            const long height = 2_500_000L;
            PowerPointTable table = slide.AddTable(2, 2, width: width, height: height);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Hello";
            cell.FillColor = "FFFF00";
            cell.Merge = (1, 2);
            table.AddRow();
            table.AddColumn();
            table.RemoveRow(2);
            table.RemoveColumn(2);
            presentation.Save();

            Console.WriteLine($"    -> Table width distributed across columns: {string.Join(", ", table.ColumnWidths)}");
            Console.WriteLine($"    -> Table height distributed across rows: {string.Join(", ", table.RowHeights)}");
        }
    }
}
