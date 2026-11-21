using System.Diagnostics;
using System.Globalization;
using OfficeIMO;
using OfficeIMO.Excel;

/// <summary>The editable <see cref="ExcelDocument"/> through its public API, as a user would write each workload.</summary>
internal static class OfficeImoLanes {
    internal static byte[] Write1M() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int column = 0; column < Workloads.LargeHeader.Length; column++) sheet.Cell(1, column + 1).SetValue(Workloads.LargeHeader[column]);
        for (int i = 1; i <= Workloads.LargeRows; i++) {
            int row = i + 1;
            sheet.Cell(row, 1).SetValue(i);
            sheet.Cell(row, 2).SetValue(Workloads.LargeName(i));
            sheet.Cell(row, 3).SetValue(Workloads.LargeAmount(i));
            sheet.Cell(row, 4).SetValue(Workloads.LargeFlag(i));
        }
        return Save(document);
    }

    internal static byte[] Styles() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int row = 1; row <= Workloads.StyleRows; row++) {
            for (int column = 1; column <= Workloads.StyleColumns; column++) {
                int style = Workloads.StyleOf(row, column);
                (byte r, byte g, byte b) = Workloads.Fill(style);
                ExcelCell cell = sheet.Cell(row, column).SetValue(row * Workloads.StyleColumns + column)
                    .SetFillColor(OfficeColor.FromRgb(r, g, b))
                    .SetFontSize(OfficeLength.Points(Workloads.FontSize(style)))
                    .SetBold(Workloads.Bold(style))
                    .SetItalic(Workloads.Italic(style));
                int border = Workloads.Border(style);
                if (border != 0) cell.SetBorder(border == 1 ? ExcelBorderStyle.Thin : ExcelBorderStyle.Medium);
            }
        }
        return Save(document);
    }

    internal static byte[] Merged() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int row = 1; row <= Workloads.MergedRows; row++) {
            sheet.Cell(row, 1).SetValue("Merged " + row.ToString(CultureInfo.InvariantCulture));
            sheet.Range(row, 1, row, 3).Merge();
        }
        return Save(document);
    }

    internal static byte[] CfDv() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int row = 1; row <= Workloads.RuleRows; row++) {
            sheet.Cell(row, 1).SetValue(row);
            sheet.Cell(row, 2).SetValue(row % 100);
            sheet.Cell(row, 3).SetValue(row % 120);
        }
        OfficeColor red = OfficeColor.FromRgb(255, 0, 0);
        for (int block = 0; block < Workloads.RuleBlocks; block++) {
            int top = block * 5 + 1;
            sheet.Range(top, 2, top + 4, 2).ConditionalFormat().CellIs(ExcelConditionalFormattingOperator.GreaterThan, "50", fillColor: red);
            sheet.Range(top, 3, top + 4, 3).Validation().WholeNumber(ExcelDataValidationOperator.Between, 1, 100);
        }
        return Save(document);
    }

    internal static byte[] NamesTables() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int table = 0; table < Workloads.Tables; table++) {
            int top = Workloads.TableTop(table);
            for (int column = 0; column < 4; column++) sheet.Cell(top, column + 1).SetValue(Workloads.TableHeader[column]);
            for (int row = 1; row <= Workloads.TableRows; row++) {
                for (int column = 1; column <= 4; column++) sheet.Cell(top + row, column).SetValue(table * 1000 + row * 4 + column);
            }
            sheet.AddTable($"A{top}:D{top + Workloads.TableRows}", hasHeader: true, name: "Table" + table.ToString(CultureInfo.InvariantCulture));
        }
        for (int name = 1; name <= Workloads.DefinedNames; name++) {
            document.SetNamedRange("Name" + name.ToString(CultureInfo.InvariantCulture), $"Data!$A${name}:$D${name}");
        }
        return Save(document);
    }

    internal static byte[] Preserve(byte[] input) {
        using var document = ExcelDocument.Load(new MemoryStream(input, writable: false));
        document.Sheets[0].Cell(Workloads.PreserveRow, Workloads.PreserveColumn).SetValue(Workloads.PreserveText);
        return Save(document);
    }

    internal static double Recalc(byte[] fixture) {
        using var document = ExcelDocument.Load(new MemoryStream(fixture, writable: false));
        ExcelSheet lookup = document.Sheets.First(sheet => sheet.Name == "Lookup");
        for (int key = 1; key <= Workloads.LookupRows; key++) lookup.Cell(key + 1, 2).SetValue(Workloads.NewLookupValue(key));
        document.Calculate();
        ExcelSheet summary = document.Sheets.First(sheet => sheet.Name == "Summary");
        double total = 0;
        for (int row = 1; row <= Workloads.Categories; row++) {
            if (!summary.TryGetCellText(row, 2, out string text)) throw new InvalidDataException("Summary!B" + row + " has no value.");
            total += double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture);
        }
        return total;
    }

    internal static (double Milliseconds, byte[] Output) AutoFit() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int row = 1; row <= Workloads.AutoFitRows; row++) {
            for (int column = 1; column <= Workloads.AutoFitColumns; column++) {
                if (Workloads.AutoFitIsText(column)) sheet.Cell(row, column).SetValue(Workloads.AutoFitText(row, column));
                else sheet.Cell(row, column).SetValue(Workloads.AutoFitNumber(row, column));
            }
        }
        var watch = Stopwatch.StartNew();
        sheet.AutoFitColumns();
        double elapsed = watch.Elapsed.TotalMilliseconds;
        return (elapsed, Save(document));
    }

    private static byte[] Save(ExcelDocument document) {
        var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }
}
