using System.Diagnostics;
using System.Globalization;
using ClosedXML.Excel;

/// <summary>ClosedXML's workbook model. XLiburLanes.cs is the same code against the XLibur fork.</summary>
internal static class ClosedXmlLanes {
    internal static byte[] Write1M() {
        using var workbook = new XLWorkbook();
        IXLWorksheet sheet = workbook.Worksheets.Add("Data");
        for (int column = 0; column < Workloads.LargeHeader.Length; column++) sheet.Cell(1, column + 1).Value = Workloads.LargeHeader[column];
        for (int i = 1; i <= Workloads.LargeRows; i++) {
            int row = i + 1;
            sheet.Cell(row, 1).Value = i;
            sheet.Cell(row, 2).Value = Workloads.LargeName(i);
            sheet.Cell(row, 3).Value = Workloads.LargeAmount(i);
            sheet.Cell(row, 4).Value = Workloads.LargeFlag(i);
        }
        return Save(workbook);
    }

    internal static byte[] Styles() {
        using var workbook = new XLWorkbook();
        IXLWorksheet sheet = workbook.Worksheets.Add("Data");
        for (int row = 1; row <= Workloads.StyleRows; row++) {
            for (int column = 1; column <= Workloads.StyleColumns; column++) {
                int style = Workloads.StyleOf(row, column);
                (byte r, byte g, byte b) = Workloads.Fill(style);
                IXLCell cell = sheet.Cell(row, column);
                cell.Value = row * Workloads.StyleColumns + column;
                IXLStyle cellStyle = cell.Style;
                cellStyle.Fill.BackgroundColor = XLColor.FromArgb(r, g, b);
                cellStyle.Font.FontSize = Workloads.FontSize(style);
                cellStyle.Font.Bold = Workloads.Bold(style);
                cellStyle.Font.Italic = Workloads.Italic(style);
                int border = Workloads.Border(style);
                if (border != 0) cellStyle.Border.OutsideBorder = border == 1 ? XLBorderStyleValues.Thin : XLBorderStyleValues.Medium;
            }
        }
        return Save(workbook);
    }

    internal static byte[] Merged() {
        using var workbook = new XLWorkbook();
        IXLWorksheet sheet = workbook.Worksheets.Add("Data");
        for (int row = 1; row <= Workloads.MergedRows; row++) {
            sheet.Cell(row, 1).Value = "Merged " + row.ToString(CultureInfo.InvariantCulture);
            sheet.Range(row, 1, row, 3).Merge();
        }
        return Save(workbook);
    }

    internal static byte[] CfDv() {
        using var workbook = new XLWorkbook();
        IXLWorksheet sheet = workbook.Worksheets.Add("Data");
        for (int row = 1; row <= Workloads.RuleRows; row++) {
            sheet.Cell(row, 1).Value = row;
            sheet.Cell(row, 2).Value = row % 100;
            sheet.Cell(row, 3).Value = row % 120;
        }
        for (int block = 0; block < Workloads.RuleBlocks; block++) {
            int top = block * 5 + 1;
            sheet.Range(top, 2, top + 4, 2).AddConditionalFormat().WhenGreaterThan(50).Fill.SetBackgroundColor(XLColor.Red);
            sheet.Range(top, 3, top + 4, 3).CreateDataValidation().WholeNumber.Between(1, 100);
        }
        return Save(workbook);
    }

    internal static byte[] NamesTables() {
        using var workbook = new XLWorkbook();
        IXLWorksheet sheet = workbook.Worksheets.Add("Data");
        for (int table = 0; table < Workloads.Tables; table++) {
            int top = Workloads.TableTop(table);
            for (int column = 0; column < 4; column++) sheet.Cell(top, column + 1).Value = Workloads.TableHeader[column];
            for (int row = 1; row <= Workloads.TableRows; row++) {
                for (int column = 1; column <= 4; column++) sheet.Cell(top + row, column).Value = table * 1000 + row * 4 + column;
            }
            sheet.Range(top, 1, top + Workloads.TableRows, 4).CreateTable("Table" + table.ToString(CultureInfo.InvariantCulture));
        }
        for (int name = 1; name <= Workloads.DefinedNames; name++) {
            workbook.DefinedNames.Add("Name" + name.ToString(CultureInfo.InvariantCulture), $"Data!$A${name}:$D${name}");
        }
        return Save(workbook);
    }

    internal static byte[] Preserve(byte[] input) {
        using var workbook = new XLWorkbook(new MemoryStream(input, writable: false));
        workbook.Worksheet(1).Cell(Workloads.PreserveRow, Workloads.PreserveColumn).Value = Workloads.PreserveText;
        return Save(workbook);
    }

    internal static double Recalc(byte[] fixture) {
        using var workbook = new XLWorkbook(new MemoryStream(fixture, writable: false));
        IXLWorksheet lookup = workbook.Worksheet("Lookup");
        for (int key = 1; key <= Workloads.LookupRows; key++) lookup.Cell(key + 1, 2).Value = Workloads.NewLookupValue(key);
        workbook.RecalculateAllFormulas();
        IXLWorksheet summary = workbook.Worksheet("Summary");
        double total = 0;
        for (int row = 1; row <= Workloads.Categories; row++) total += summary.Cell(row, 2).Value.GetNumber();
        return total;
    }

    internal static (double Milliseconds, byte[] Output) AutoFit() {
        using var workbook = new XLWorkbook();
        IXLWorksheet sheet = workbook.Worksheets.Add("Data");
        for (int row = 1; row <= Workloads.AutoFitRows; row++) {
            for (int column = 1; column <= Workloads.AutoFitColumns; column++) {
                if (Workloads.AutoFitIsText(column)) sheet.Cell(row, column).Value = Workloads.AutoFitText(row, column);
                else sheet.Cell(row, column).Value = Workloads.AutoFitNumber(row, column);
            }
        }
        var watch = Stopwatch.StartNew();
        sheet.Columns(1, Workloads.AutoFitColumns).AdjustToContents();
        double elapsed = watch.Elapsed.TotalMilliseconds;
        return (elapsed, Save(workbook));
    }

    private static byte[] Save(XLWorkbook workbook) {
        var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }
}
