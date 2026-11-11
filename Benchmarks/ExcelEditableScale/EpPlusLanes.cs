using System.Diagnostics;
using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;

/// <summary>EPPlus's workbook model, declared for non-commercial use as EPPlus 8 requires.</summary>
internal static class EpPlusLanes {
    static EpPlusLanes() => ExcelPackage.License.SetNonCommercialOrganization("OfficeIMO benchmarks");

    internal static byte[] Write1M() {
        using var package = new ExcelPackage();
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        for (int column = 0; column < Workloads.LargeHeader.Length; column++) sheet.Cells[1, column + 1].Value = Workloads.LargeHeader[column];
        for (int i = 1; i <= Workloads.LargeRows; i++) {
            int row = i + 1;
            sheet.Cells[row, 1].Value = i;
            sheet.Cells[row, 2].Value = Workloads.LargeName(i);
            sheet.Cells[row, 3].Value = Workloads.LargeAmount(i);
            sheet.Cells[row, 4].Value = Workloads.LargeFlag(i);
        }
        return package.GetAsByteArray();
    }

    internal static byte[] Styles() {
        using var package = new ExcelPackage();
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        for (int row = 1; row <= Workloads.StyleRows; row++) {
            for (int column = 1; column <= Workloads.StyleColumns; column++) {
                int style = Workloads.StyleOf(row, column);
                (byte r, byte g, byte b) = Workloads.Fill(style);
                ExcelRange cell = sheet.Cells[row, column];
                cell.Value = row * Workloads.StyleColumns + column;
                ExcelStyle cellStyle = cell.Style;
                cellStyle.Fill.PatternType = ExcelFillStyle.Solid;
                cellStyle.Fill.BackgroundColor.SetColor(255, r, g, b);
                cellStyle.Font.Size = (float)Workloads.FontSize(style);
                cellStyle.Font.Bold = Workloads.Bold(style);
                cellStyle.Font.Italic = Workloads.Italic(style);
                int border = Workloads.Border(style);
                if (border != 0) cellStyle.Border.BorderAround(border == 1 ? ExcelBorderStyle.Thin : ExcelBorderStyle.Medium);
            }
        }
        return package.GetAsByteArray();
    }

    internal static byte[] Merged() {
        using var package = new ExcelPackage();
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        for (int row = 1; row <= Workloads.MergedRows; row++) {
            sheet.Cells[row, 1].Value = "Merged " + row.ToString(CultureInfo.InvariantCulture);
            sheet.Cells[row, 1, row, 3].Merge = true;
        }
        return package.GetAsByteArray();
    }

    internal static byte[] CfDv() {
        using var package = new ExcelPackage();
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        for (int row = 1; row <= Workloads.RuleRows; row++) {
            sheet.Cells[row, 1].Value = row;
            sheet.Cells[row, 2].Value = row % 100;
            sheet.Cells[row, 3].Value = row % 120;
        }
        for (int block = 0; block < Workloads.RuleBlocks; block++) {
            int top = block * 5 + 1;
            var rule = sheet.ConditionalFormatting.AddGreaterThan(sheet.Cells[top, 2, top + 4, 2]);
            rule.Formula = "50";
            rule.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule.Style.Fill.BackgroundColor.Color = System.Drawing.Color.Red;
            var validation = sheet.DataValidations.AddIntegerValidation(sheet.Cells[top, 3, top + 4, 3].Address);
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Formula.Value = 1;
            validation.Formula2.Value = 100;
        }
        return package.GetAsByteArray();
    }

    internal static byte[] NamesTables() {
        using var package = new ExcelPackage();
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        for (int table = 0; table < Workloads.Tables; table++) {
            int top = Workloads.TableTop(table);
            for (int column = 0; column < 4; column++) sheet.Cells[top, column + 1].Value = Workloads.TableHeader[column];
            for (int row = 1; row <= Workloads.TableRows; row++) {
                for (int column = 1; column <= 4; column++) sheet.Cells[top + row, column].Value = table * 1000 + row * 4 + column;
            }
            sheet.Tables.Add(sheet.Cells[top, 1, top + Workloads.TableRows, 4], "Table" + table.ToString(CultureInfo.InvariantCulture));
        }
        for (int name = 1; name <= Workloads.DefinedNames; name++) {
            package.Workbook.Names.Add("Name" + name.ToString(CultureInfo.InvariantCulture), sheet.Cells[name, 1, name, 4]);
        }
        return package.GetAsByteArray();
    }

    internal static byte[] Preserve(byte[] input) {
        using var package = new ExcelPackage(new MemoryStream(input, writable: false));
        package.Workbook.Worksheets[0].Cells[Workloads.PreserveRow, Workloads.PreserveColumn].Value = Workloads.PreserveText;
        return package.GetAsByteArray();
    }

    internal static double Recalc(byte[] fixture) {
        using var package = new ExcelPackage(new MemoryStream(fixture, writable: false));
        ExcelWorksheet lookup = package.Workbook.Worksheets["Lookup"];
        for (int key = 1; key <= Workloads.LookupRows; key++) lookup.Cells[key + 1, 2].Value = Workloads.NewLookupValue(key);
        package.Workbook.Calculate();
        ExcelWorksheet summary = package.Workbook.Worksheets["Summary"];
        double total = 0;
        for (int row = 1; row <= Workloads.Categories; row++) total += Convert.ToDouble(summary.Cells[row, 2].Value, CultureInfo.InvariantCulture);
        return total;
    }

    internal static (double Milliseconds, byte[] Output) AutoFit() {
        using var package = new ExcelPackage();
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        for (int row = 1; row <= Workloads.AutoFitRows; row++) {
            for (int column = 1; column <= Workloads.AutoFitColumns; column++) {
                if (Workloads.AutoFitIsText(column)) sheet.Cells[row, column].Value = Workloads.AutoFitText(row, column);
                else sheet.Cells[row, column].Value = Workloads.AutoFitNumber(row, column);
            }
        }
        var watch = Stopwatch.StartNew();
        sheet.Cells[1, 1, Workloads.AutoFitRows, Workloads.AutoFitColumns].AutoFitColumns();
        double elapsed = watch.Elapsed.TotalMilliseconds;
        return (elapsed, package.GetAsByteArray());
    }
}
