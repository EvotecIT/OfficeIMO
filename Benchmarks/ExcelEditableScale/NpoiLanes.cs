using System.Diagnostics;
using System.Globalization;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

/// <summary>
/// NPOI's editable XSSF model. NPOI styles are workbook objects, so the styles lane creates each of the 2,000 styles
/// once and assigns them, which is how NPOI code has to be written. Merges use AddMergedRegionUnsafe because the
/// checked AddMergedRegion compares each new region with every earlier one, which does not finish at this scale.
/// </summary>
internal static class NpoiLanes {
    internal static byte[] Write1M() {
        using var workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Data");
        IRow header = sheet.CreateRow(0);
        for (int column = 0; column < Workloads.LargeHeader.Length; column++) header.CreateCell(column).SetCellValue(Workloads.LargeHeader[column]);
        for (int i = 1; i <= Workloads.LargeRows; i++) {
            IRow row = sheet.CreateRow(i);
            row.CreateCell(0).SetCellValue(i);
            row.CreateCell(1).SetCellValue(Workloads.LargeName(i));
            row.CreateCell(2).SetCellValue(Workloads.LargeAmount(i));
            row.CreateCell(3).SetCellValue(Workloads.LargeFlag(i));
        }
        return Save(workbook);
    }

    internal static byte[] Styles() {
        using var workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Data");
        var styles = new ICellStyle[Workloads.StyleCount];
        for (int style = 0; style < Workloads.StyleCount; style++) {
            (byte r, byte g, byte b) = Workloads.Fill(style);
            var cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = Workloads.FontSize(style);
            font.IsBold = Workloads.Bold(style);
            font.IsItalic = Workloads.Italic(style);
            cellStyle.SetFont(font);
            var fill = new XSSFColor();
            fill.SetRgb(new[] { r, g, b });
            cellStyle.SetFillForegroundColor(fill);
            cellStyle.FillPattern = FillPattern.SolidForeground;
            int border = Workloads.Border(style);
            if (border != 0) {
                BorderStyle line = border == 1 ? BorderStyle.Thin : BorderStyle.Medium;
                cellStyle.BorderTop = line;
                cellStyle.BorderBottom = line;
                cellStyle.BorderLeft = line;
                cellStyle.BorderRight = line;
            }
            styles[style] = cellStyle;
        }
        for (int row = 1; row <= Workloads.StyleRows; row++) {
            IRow sheetRow = sheet.CreateRow(row - 1);
            for (int column = 1; column <= Workloads.StyleColumns; column++) {
                ICell cell = sheetRow.CreateCell(column - 1);
                cell.SetCellValue(row * Workloads.StyleColumns + column);
                cell.CellStyle = styles[Workloads.StyleOf(row, column)];
            }
        }
        return Save(workbook);
    }

    internal static byte[] Merged() {
        using var workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Data");
        for (int row = 1; row <= Workloads.MergedRows; row++) {
            sheet.CreateRow(row - 1).CreateCell(0).SetCellValue("Merged " + row.ToString(CultureInfo.InvariantCulture));
            sheet.AddMergedRegionUnsafe(new CellRangeAddress(row - 1, row - 1, 0, 2));
        }
        return Save(workbook);
    }

    internal static byte[] CfDv() {
        using var workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Data");
        for (int row = 1; row <= Workloads.RuleRows; row++) {
            IRow sheetRow = sheet.CreateRow(row - 1);
            sheetRow.CreateCell(0).SetCellValue(row);
            sheetRow.CreateCell(1).SetCellValue(row % 100);
            sheetRow.CreateCell(2).SetCellValue(row % 120);
        }
        ISheetConditionalFormatting formatting = sheet.SheetConditionalFormatting;
        IDataValidationHelper helper = sheet.GetDataValidationHelper();
        for (int block = 0; block < Workloads.RuleBlocks; block++) {
            int top = block * 5;
            IConditionalFormattingRule rule = formatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThan, "50");
            IPatternFormatting fill = rule.CreatePatternFormatting();
            fill.FillBackgroundColor = IndexedColors.Red.Index;
            fill.FillPattern = FillPattern.SolidForeground;
            formatting.AddConditionalFormatting(new[] { new CellRangeAddress(top, top + 4, 1, 1) }, rule);
            IDataValidationConstraint constraint = helper.CreateintConstraint(OperatorType.BETWEEN, "1", "100");
            sheet.AddValidationData(helper.CreateValidation(constraint, new CellRangeAddressList(top, top + 4, 2, 2)));
        }
        return Save(workbook);
    }

    internal static byte[] NamesTables() {
        using var workbook = new XSSFWorkbook();
        var sheet = (XSSFSheet)workbook.CreateSheet("Data");
        for (int table = 0; table < Workloads.Tables; table++) {
            int top = Workloads.TableTop(table) - 1;
            IRow header = sheet.CreateRow(top);
            for (int column = 0; column < 4; column++) header.CreateCell(column).SetCellValue(Workloads.TableHeader[column]);
            for (int row = 1; row <= Workloads.TableRows; row++) {
                IRow sheetRow = sheet.CreateRow(top + row);
                for (int column = 1; column <= 4; column++) sheetRow.CreateCell(column - 1).SetCellValue(table * 1000 + row * 4 + column);
            }
            XSSFTable created = sheet.CreateTable();
            created.CellReferences = new AreaReference(new CellReference(top, 0), new CellReference(top + Workloads.TableRows, 3), workbook.SpreadsheetVersion);
            created.UpdateHeaders();
            created.Name = "Table" + table.ToString(CultureInfo.InvariantCulture);
            created.DisplayName = created.Name;
        }
        for (int name = 1; name <= Workloads.DefinedNames; name++) {
            IName definedName = workbook.CreateName();
            definedName.NameName = "Name" + name.ToString(CultureInfo.InvariantCulture);
            definedName.RefersToFormula = $"Data!$A${name}:$D${name}";
        }
        return Save(workbook);
    }

    internal static byte[] Preserve(byte[] input) {
        using var workbook = new XSSFWorkbook(new MemoryStream(input, writable: false));
        ISheet sheet = workbook.GetSheetAt(0);
        IRow row = sheet.GetRow(Workloads.PreserveRow - 1) ?? sheet.CreateRow(Workloads.PreserveRow - 1);
        (row.GetCell(Workloads.PreserveColumn - 1) ?? row.CreateCell(Workloads.PreserveColumn - 1)).SetCellValue(Workloads.PreserveText);
        return Save(workbook);
    }

    internal static double Recalc(byte[] fixture) {
        using var workbook = new XSSFWorkbook(new MemoryStream(fixture, writable: false));
        ISheet lookup = workbook.GetSheet("Lookup");
        for (int key = 1; key <= Workloads.LookupRows; key++) {
            IRow row = lookup.GetRow(key) ?? lookup.CreateRow(key);
            (row.GetCell(1) ?? row.CreateCell(1)).SetCellValue(Workloads.NewLookupValue(key));
        }
        XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
        ISheet summary = workbook.GetSheet("Summary");
        double total = 0;
        for (int row = 0; row < Workloads.Categories; row++) total += summary.GetRow(row).GetCell(1).NumericCellValue;
        return total;
    }

    internal static (double Milliseconds, byte[] Output) AutoFit() {
        using var workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Data");
        for (int row = 1; row <= Workloads.AutoFitRows; row++) {
            IRow sheetRow = sheet.CreateRow(row - 1);
            for (int column = 1; column <= Workloads.AutoFitColumns; column++) {
                if (Workloads.AutoFitIsText(column)) sheetRow.CreateCell(column - 1).SetCellValue(Workloads.AutoFitText(row, column));
                else sheetRow.CreateCell(column - 1).SetCellValue(Workloads.AutoFitNumber(row, column));
            }
        }
        var watch = Stopwatch.StartNew();
        for (int column = 0; column < Workloads.AutoFitColumns; column++) sheet.AutoSizeColumn(column);
        double elapsed = watch.Elapsed.TotalMilliseconds;
        return (elapsed, Save(workbook));
    }

    private static byte[] Save(XSSFWorkbook workbook) {
        var output = new MemoryStream();
        workbook.Write(output, leaveOpen: true);
        return output.ToArray();
    }
}
