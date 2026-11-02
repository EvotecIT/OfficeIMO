using System.Globalization;
using System.Text;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Deterministic data and source workbooks with the shapes of the public XLBench scenarios.</summary>
/// <remarks>ClosedXML is the neutral producer of the source workbooks, as in XLBench. All setup runs outside measurement.</remarks>
internal static class XlBenchShapeData {
    internal const int WriteRows = 50_000;
    internal const int ReadRows = 50_000;
    internal const int EditRows = 500;
    internal const int EditColumns = 20;
    internal const int EditSumColumn = EditColumns + 1;
    internal const double EditColumnAValue = 20d;
    internal const int InsertAt = 2;
    internal const int InsertCount = 2;
    internal const double InsertedValue = 10d;
    internal const double EditedValue = 7d;
    internal const string EditedText = "Edited";
    internal const int SheetCount = 20;
    internal const int RowsPerSheet = 500;
    internal const int SheetColumns = 12;
    internal const int StockWeeks = 260;
    internal const int StockSymbols = 20;
    internal const int TemplateRows = 2_000;
    internal const int TemplateFirstRow = 6;
    internal const int PropertyRows = 1_000;
    internal const int PropertyColumns = 8;
    internal const int WideRows = 10_000;
    internal const int WideColumns = 150;
    internal const string ReportSheet = "Weekly Prices";
    internal const string ReportTitle = "Weekly closing price by symbol";
    internal const string DateFormat = "yyyy-mm-dd";
    internal const string PriceFormat = "#,##0.00";
    internal const string PropertyTitle = "XLBench properties round trip";
    internal const string PropertyCategory = "Benchmark";
    internal const string TemplateTitle = "Sales Report — January 2026";
    internal const string TemplateRegion = "EMEA";
    internal static readonly DateTime TemplateDate = new(2026, 1, 31);

    internal static readonly string[] WriteStrings = new string[WriteRows];
    internal static readonly double[] WriteNumbers = new double[WriteRows];
    internal static readonly DateTime[] WriteDates = new DateTime[WriteRows];
    internal static readonly string[] Symbols = BuildSymbols();
    internal static readonly DateTime[] WeekEndings = new DateTime[StockWeeks];
    internal static readonly double[][] Prices = new double[StockWeeks][];
    internal static readonly double[][] EditValues = BuildEditValues();
    internal static readonly int[] RowsToDelete = BuildRowsToDelete();
    internal static readonly int[] SurvivingRows = BuildSurvivingRows();
    internal static readonly TemplateLine[] Lines = BuildLines();
    internal static readonly string ChartPriceRange = $"C3:{ColumnName(2 + StockSymbols)}{1 + StockWeeks}";

    private static readonly Lazy<byte[]> LazyRead = new(BuildRead);
    private static readonly Lazy<byte[]> LazyEdit = new(BuildEdit);
    private static readonly Lazy<byte[]> LazyProperties = new(BuildProperties);
    private static readonly Lazy<byte[]> LazySheets = new(BuildSheets);
    private static readonly Lazy<byte[]> LazyTemplate = new(LoadTemplate);
    private static readonly Lazy<byte[]> LazyWide = new(BuildWide);

    internal static byte[] ReadWorkbook => LazyRead.Value;
    internal static byte[] EditWorkbook => LazyEdit.Value;
    internal static byte[] PropertiesWorkbook => LazyProperties.Value;
    internal static byte[] SheetsWorkbook => LazySheets.Value;
    internal static byte[] TemplateWorkbook => LazyTemplate.Value;
    internal static byte[] WideWorkbook => LazyWide.Value;
    internal static long SheetsCellCount => (long)SheetCount * (RowsPerSheet + 1) * SheetColumns;
    internal static int EditLastRowAfterDelete => 1 + SurvivingRows.Length;

    internal readonly record struct TemplateLine(DateTime Date, string Product, string Category, int Quantity, double UnitPrice);

    static XlBenchShapeData() {
        var random = new Random(42);
        var start = new DateTime(2020, 1, 1);
        for (int i = 0; i < WriteRows; i++) {
            WriteStrings[i] = $"Item {i} - {random.Next(1000):D4}";
            WriteNumbers[i] = Math.Round(random.NextDouble() * 10000, 2);
            WriteDates[i] = start.AddDays(random.Next(0, 1500));
        }

        var week = new DateTime(2021, 8, 6);
        for (int w = 0; w < StockWeeks; w++) {
            WeekEndings[w] = week.AddDays(w * 7);
            Prices[w] = new double[StockSymbols];
            for (int s = 0; s < StockSymbols; s++) {
                Prices[w][s] = Math.Round(40 + random.NextDouble() * 600, 2);
            }
        }
    }

    internal static string ColumnName(int column) {
        var builder = new StringBuilder();
        for (; column > 0; column = (column - 1) / 26) {
            builder.Insert(0, (char)('A' + (column - 1) % 26));
        }
        return builder.ToString();
    }

    internal static string EditFormula(int row) => $"SUM(A{row}:{ColumnName(EditColumns)}{row})";

    internal static string InsertFormula(int row) => $"SUM(A{row}:{ColumnName(EditColumns + InsertCount)}{row})";

    internal static double EditExpectedTotal(int survivingIndex) {
        double[] source = EditValues[SurvivingRows[survivingIndex]];
        double total = EditColumnAValue;
        for (int column = 1; column < source.Length; column++) {
            total += source[column];
        }
        return total;
    }

    internal static double InsertExpectedTotal(int rowIndex) {
        double total = InsertCount * InsertedValue;
        foreach (double value in EditValues[rowIndex]) {
            total += value;
        }
        return total;
    }

    internal static (int Row, int Column) ParseCell(string reference) {
        int bang = reference.LastIndexOf('!');
        int row = 0;
        int column = 0;
        foreach (char ch in bang < 0 ? reference : reference.Substring(bang + 1)) {
            if (char.IsAsciiLetterUpper(ch)) {
                column = column * 26 + ch - 'A' + 1;
            } else if (char.IsAsciiDigit(ch)) {
                row = row * 10 + ch - '0';
            }
        }
        return (row, column);
    }

    private static string[] BuildSymbols() {
        var symbols = new string[StockSymbols];
        for (int i = 0; i < symbols.Length; i++) {
            symbols[i] = "SYM" + (char)('A' + i);
        }
        return symbols;
    }

    private static double[][] BuildEditValues() {
        var random = new Random(4000);
        var rows = new double[EditRows][];
        for (int row = 0; row < rows.Length; row++) {
            rows[row] = new double[EditColumns];
            for (int column = 0; column < EditColumns; column++) {
                rows[row][column] = (row + column) % 2 == 0
                    ? random.Next(-99_999, 99_999)
                    : Math.Round((random.NextDouble() - 0.5) * 200_000, 4);
            }
        }
        return rows;
    }

    private static int[] BuildRowsToDelete() {
        var rows = new List<int>();
        for (int i = 0; i < EditRows; i++) {
            if ((i + 1) % 3 == 0) {
                rows.Add(2 + i);
            }
        }
        return rows.ToArray();
    }

    private static int[] BuildSurvivingRows() {
        var rows = new List<int>();
        for (int i = 0; i < EditRows; i++) {
            if ((i + 1) % 3 != 0) {
                rows.Add(i);
            }
        }
        return rows.ToArray();
    }

    private static TemplateLine[] BuildLines() {
        (string Product, string Category, double Price)[] catalogue = [
            ("Laptop 14\"", "Hardware", 1249.00), ("Laptop 16\"", "Hardware", 1799.00), ("Monitor 27\"", "Hardware", 329.99),
            ("Docking station", "Hardware", 189.50), ("Office suite", "Software", 99.00), ("Antivirus", "Software", 39.95),
            ("Design suite", "Software", 649.00), ("Backup service", "Software", 7.50), ("Installation", "Services", 150.00),
            ("Support hour", "Services", 85.00), ("Data migration", "Services", 1200.00), ("Keyboard", "Accessories", 49.99),
            ("Mouse", "Accessories", 24.99), ("Headset", "Accessories", 79.00), ("USB-C cable", "Accessories", 12.49),
            ("Onboarding course", "Training", 299.00), ("Advanced workshop", "Training", 890.00), ("E-learning seat", "Training", 45.00)
        ];
        var random = new Random(20260131);
        var start = new DateTime(2026, 1, 1);
        var lines = new TemplateLine[TemplateRows];
        for (int i = 0; i < lines.Length; i++) {
            var (product, category, price) = catalogue[random.Next(catalogue.Length)];
            lines[i] = new TemplateLine(start.AddDays(i * 31 / TemplateRows), product, category, random.Next(1, 40), price);
        }
        return lines;
    }

    private static byte[] Save(XLWorkbook workbook) {
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static byte[] BuildRead() {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Data");
        var random = new Random(42);
        var start = new DateTime(2020, 1, 1);
        string[] regions = ["North", "South", "East", "West", "Central"];
        string[] statuses = ["Active", "Pending", "Closed", "Review", "Draft"];
        for (int i = 0; i < ReadRows; i++) {
            int row = i + 1;
            int seed = random.Next(10000);
            sheet.Cell(row, 1).Value = $"Item {i}-{seed}";
            sheet.Cell(row, 2).Value = $"Cat-{i % 12}";
            sheet.Cell(row, 3).Value = regions[i % regions.Length];
            sheet.Cell(row, 4).Value = Math.Round(random.NextDouble() * 10000, 2);
            sheet.Cell(row, 5).Value = random.Next(1, 5000);
            sheet.Cell(row, 6).Value = Math.Round(random.NextDouble(), 4);
            sheet.Cell(row, 7).Value = Math.Round(random.NextDouble() * 1000, 2);
            sheet.Cell(row, 8).Value = random.Next(0, 100);
            sheet.Cell(row, 9).Value = start.AddDays(random.Next(0, 2000));
            sheet.Cell(row, 10).Value = start.AddDays(random.Next(0, 2000));
            sheet.Cell(row, 11).Value = start.AddDays(random.Next(0, 2000));
            sheet.Cell(row, 12).Value = statuses[i % statuses.Length];
            sheet.Cell(row, 13).Value = $"Note for row {row} with seed {seed}";
            sheet.Cell(row, 14).Value = $"CODE-{seed:D5}";
            sheet.Cell(row, 15).FormulaA1 = $"SUM(D{row}:H{row})";
        }
        return Save(workbook);
    }

    private static byte[] BuildWide() {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Wide");
        for (int column = 1; column <= WideColumns; column++) {
            sheet.Cell(1, column).Value = $"Column_{column}";
        }
        var random = new Random(42);
        for (int row = 2; row <= WideRows; row++) {
            for (int column = 1; column <= WideColumns; column++) {
                if (column % 5 == 0) {
                    sheet.Cell(row, column).Value = $"Text {random.Next(5000)}";
                } else {
                    sheet.Cell(row, column).Value = Math.Round(random.NextDouble() * 10000, 2);
                }
            }
        }
        return Save(workbook);
    }

    private static byte[] BuildEdit() {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Numbers");
        for (int column = 1; column <= EditColumns; column++) {
            sheet.Cell(1, column).Value = string.Create(CultureInfo.InvariantCulture, $"Column_{column:D2}");
        }
        sheet.Cell(1, EditSumColumn).Value = "Total";
        for (int i = 0; i < EditRows; i++) {
            int row = 2 + i;
            for (int column = 0; column < EditColumns; column++) {
                sheet.Cell(row, column + 1).Value = EditValues[i][column];
            }
            sheet.Cell(row, EditSumColumn).FormulaA1 = EditFormula(row);
            if (row % 2 == 0) {
                sheet.Range(row, 1, row, EditSumColumn).Style.Font.Bold = true;
            }
        }
        return Save(workbook);
    }

    private static byte[] BuildProperties() {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Numbers");
        for (int column = 1; column <= PropertyColumns; column++) {
            sheet.Cell(1, column).Value = $"Column_{column}";
        }
        var random = new Random(42);
        for (int i = 0; i < PropertyRows; i++) {
            int row = 2 + i;
            sheet.Cell(row, 1).Value = Math.Round(random.NextDouble() * 10000, 2);
            sheet.Cell(row, 2).Value = random.Next(1, 5000);
            sheet.Cell(row, 3).Value = Math.Round(random.NextDouble(), 4);
            sheet.Cell(row, 4).Value = Math.Round(random.NextDouble() * 1000, 2);
            sheet.Cell(row, 5).Value = random.Next(0, 100);
            sheet.Cell(row, 6).Value = Math.Round(random.NextDouble() * -5000, 3);
            sheet.Cell(row, 7).Value = random.Next(-1000, 1000);
            sheet.Cell(row, 8).Value = Math.Round(random.NextDouble() * 100, 4);
        }
        return Save(workbook);
    }

    private static byte[] BuildSheets() {
        const int vocabularySize = 40_000;
        const int paletteSize = 1_200;
        var random = new Random(42);
        string[] syllables = [
            "ka", "lo", "mi", "ra", "sen", "tor", "vel", "nu", "pa", "dri", "qua", "zen", "ho", "fi", "gar", "bel", "co", "da", "esh", "ju",
            "lin", "mor", "ne", "os", "pri", "rek", "sa", "tu", "vin", "wy", "xan", "yo", "zul", "ar", "bo", "cy", "del", "en", "fa", "gu"
        ];
        var vocabulary = new HashSet<string>(vocabularySize, StringComparer.Ordinal);
        var phrase = new StringBuilder();
        while (vocabulary.Count < vocabularySize) {
            phrase.Clear();
            int words = random.Next(2, 6);
            for (int word = 0; word < words; word++) {
                if (word > 0) {
                    phrase.Append(' ');
                }
                phrase.Append(syllables[random.Next(syllables.Length)]).Append(syllables[random.Next(syllables.Length)]);
                if (random.Next(3) == 0) {
                    phrase.Append(syllables[random.Next(syllables.Length)]);
                }
            }
            vocabulary.Add(phrase.ToString());
        }
        string[] phrases = vocabulary.ToArray();
        string[] headers = ["Name", "Owner", "Team", "Region", "Product", "Status", "Channel", "Tags", "Notes", "Amount", "Due", "Rate"];
        string[] amountFormats = ["#,##0.00", "0.00", "#,##0", "[$$-409]#,##0.00"];
        string[] dateFormats = ["yyyy-mm-dd", "dd mmm yyyy", "d/m/yy", "mmm-yy"];
        string[] rateFormats = ["0.00%", "0.0%", "0.000", "0.00E+00"];
        XLColor[] fontColors = [XLColor.Black, XLColor.DarkBlue, XLColor.DarkRed, XLColor.DarkGreen, XLColor.Purple, XLColor.SaddleBrown];
        XLColor[] fillColors = [XLColor.NoColor, XLColor.LightYellow, XLColor.LightBlue, XLColor.LightGreen, XLColor.LightPink, XLColor.Lavender, XLColor.Beige, XLColor.Honeydew, XLColor.MistyRose, XLColor.AliceBlue];
        XLBorderStyleValues[] borders = [XLBorderStyleValues.None, XLBorderStyleValues.Thin, XLBorderStyleValues.Dotted, XLBorderStyleValues.Dashed, XLBorderStyleValues.Medium];
        var start = new DateTime(2020, 1, 1);

        using var workbook = new XLWorkbook();
        var paletteSheet = workbook.AddWorksheet("Palette");
        var palette = new IXLStyle[paletteSize];
        for (int p = 0; p < paletteSize; p++) {
            var style = paletteSheet.Cell(p + 1, 1).Style;
            style.Font.FontColor = fontColors[p % fontColors.Length];
            style.Font.Bold = p / 6 % 2 == 1;
            style.Font.Italic = p / 12 % 2 == 1;
            var fill = fillColors[p / 24 % fillColors.Length];
            if (fill != XLColor.NoColor) {
                style.Fill.BackgroundColor = fill;
            }
            var border = borders[p / 240 % borders.Length];
            style.Border.LeftBorder = border;
            style.Border.RightBorder = border;
            style.Border.TopBorder = border;
            style.Border.BottomBorder = border;
            palette[p] = style;
        }

        for (int s = 0; s < SheetCount; s++) {
            var sheet = workbook.AddWorksheet($"Data {s + 1:D2}");
            for (int column = 0; column < SheetColumns; column++) {
                sheet.Cell(1, column + 1).Value = headers[column];
                sheet.Cell(1, column + 1).Style.Font.Bold = true;
            }
            for (int row = 2; row <= RowsPerSheet + 1; row++) {
                for (int column = 1; column <= 9; column++) {
                    var cell = sheet.Cell(row, column);
                    string text = random.NextDouble() < 0.3 ? phrases[random.Next(250)] : phrases[random.Next(phrases.Length)];
                    if (random.Next(10) == 0) {
                        string[] parts = text.Split(' ');
                        var rich = cell.CreateRichText();
                        for (int part = 0; part < parts.Length; part++) {
                            var run = rich.AddText(part + 1 < parts.Length ? parts[part] + " " : parts[part]);
                            switch (random.Next(4)) {
                                case 0: run.SetBold(); break;
                                case 1: run.SetItalic(); break;
                                case 2: run.SetFontColor(XLColor.Red); break;
                                default: run.SetFontSize(14).SetUnderline(); break;
                            }
                        }
                    } else {
                        cell.Value = text;
                    }
                    cell.Style = palette[random.Next(paletteSize)];
                }

                var amount = sheet.Cell(row, 10);
                amount.Value = Math.Round(random.NextDouble() * 20_000 - 2_000, 2);
                amount.Style = palette[random.Next(100)];
                amount.Style.NumberFormat.Format = amountFormats[random.Next(amountFormats.Length)];
                var due = sheet.Cell(row, 11);
                due.Value = start.AddDays(random.Next(0, 2_000));
                due.Style = palette[random.Next(100)];
                due.Style.NumberFormat.Format = dateFormats[random.Next(dateFormats.Length)];
                var rate = sheet.Cell(row, 12);
                rate.Value = Math.Round(random.NextDouble() * 0.5, 4);
                rate.Style = palette[random.Next(100)];
                rate.Style.NumberFormat.Format = rateFormats[random.Next(rateFormats.Length)];
            }
        }
        paletteSheet.Delete();
        return Save(workbook);
    }

    private static byte[] LoadTemplate() {
        using Stream stream = typeof(XlBenchShapeData).Assembly.GetManifestResourceStream("XlBenchShapes.sales_template.xlsx")
            ?? throw new InvalidOperationException("The embedded sales template is missing.");
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
}
