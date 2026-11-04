using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>The XLBench calc shape: a price list, order lines that look prices up, and a summary of conditional aggregates.</summary>
internal static partial class XlBenchShapeData {
    internal const string CalcProductsSheet = "Products";
    internal const string CalcOrdersSheet = "Orders";
    internal const string CalcSummarySheet = "Summary";
    internal const int CalcProducts = 200;
    internal const int CalcOrders = 1_000;
    internal const int CalcUnknownEveryNth = 97;
    internal const int CalcDiscountQuantity = 50;
    internal const double CalcPriceIncrease = 3d;
    internal const int CalcUnitPriceColumn = 3;
    internal const int CalcStatColumn = 2;
    internal const string CalcUnknownCategory = "Unknown";
    internal static readonly string[] CalcCategories = ["Hardware", "Software", "Services", "Networking", "Storage", "Peripherals", "Licensing", "Training", CalcUnknownCategory];
    internal static readonly string[] CalcRegions = ["North", "South", "East", "West"];
    internal static int CalcLastProductRow => 1 + CalcProducts;
    internal static int CalcLastOrderRow => 1 + CalcOrders;
    internal static int CalcGrandTotalRow => 1 + CalcCategories.Length + 3;

    internal readonly record struct CalcProduct(string Id, string Category, double UnitPrice, double DiscountRate);

    internal readonly record struct CalcOrder(string ProductId, int Quantity, string Region);

    internal static readonly CalcProduct[] CalcProductRows = new CalcProduct[CalcProducts];
    internal static readonly CalcOrder[] CalcOrderRows = new CalcOrder[CalcOrders];
    private static readonly Lazy<byte[]> LazyCalc = new(BuildCalc);

    internal static byte[] CalcWorkbook => LazyCalc.Value;

    internal static double CalcNewPrice(int product) => CalcProductRows[product].UnitPrice + CalcPriceIncrease;

    private static void BuildCalcRows() {
        string[] categories = CalcCategories[..^1];
        double[] rates = [0.05, 0.10, 0.15, 0.20];
        var random = new Random(20261020);
        for (int i = 0; i < CalcProducts; i++) {
            CalcProductRows[i] = new CalcProduct($"P{i + 1:000}", categories[random.Next(categories.Length)], random.Next(5, 501), rates[random.Next(rates.Length)]);
        }
        for (int i = 0; i < CalcOrders; i++) {
            string product = (i + 1) % CalcUnknownEveryNth == 0 ? "P999" : CalcProductRows[random.Next(CalcProducts)].Id;
            CalcOrderRows[i] = new CalcOrder(product, random.Next(1, 101), CalcRegions[random.Next(CalcRegions.Length)]);
        }
    }

    /// <summary>Each order's net and category, and the grand total, after the price increase, computed without an engine.</summary>
    internal static (double[] Net, string[] Category, double GrandTotal) CalcExpected() {
        var byId = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < CalcProducts; i++) {
            byId[CalcProductRows[i].Id] = i;
        }
        var net = new double[CalcOrders];
        var category = new string[CalcOrders];
        double total = 0;
        for (int i = 0; i < CalcOrders; i++) {
            CalcOrder order = CalcOrderRows[i];
            double price = 0, rate = 0;
            category[i] = CalcUnknownCategory;
            if (byId.TryGetValue(order.ProductId, out int product)) {
                price = CalcNewPrice(product);
                rate = CalcProductRows[product].DiscountRate;
                category[i] = CalcProductRows[product].Category;
            }
            double gross = order.Quantity * price;
            net[i] = Math.Round(gross - (order.Quantity >= CalcDiscountQuantity ? gross * rate : 0d), 2, MidpointRounding.AwayFromZero);
            total += net[i];
        }
        return (net, category, total);
    }

    private static byte[] BuildCalc() {
        string productTable = $"{CalcProductsSheet}!$A$2:$D${CalcLastProductRow}";
        string ProductColumn(string letter) => $"{CalcProductsSheet}!${letter}$2:${letter}${CalcLastProductRow}";
        string OrderColumn(string letter) => $"{CalcOrdersSheet}!${letter}$2:${letter}${CalcLastOrderRow}";

        using var workbook = new XLWorkbook();
        var products = workbook.AddWorksheet(CalcProductsSheet);
        string[] productHeaders = ["ProductId", "Category", "UnitPrice", "DiscountRate"];
        for (int c = 0; c < productHeaders.Length; c++) {
            products.Cell(1, c + 1).Value = productHeaders[c];
        }
        for (int i = 0; i < CalcProducts; i++) {
            CalcProduct product = CalcProductRows[i];
            products.Cell(i + 2, 1).Value = product.Id;
            products.Cell(i + 2, 2).Value = product.Category;
            products.Cell(i + 2, 3).Value = product.UnitPrice;
            products.Cell(i + 2, 4).Value = product.DiscountRate;
        }

        var orders = workbook.AddWorksheet(CalcOrdersSheet);
        string[] orderHeaders = ["OrderId", "ProductId", "Quantity", "Region", "UnitPrice", "Category", "Gross", "Discount", "Net"];
        for (int c = 0; c < orderHeaders.Length; c++) {
            orders.Cell(1, c + 1).Value = orderHeaders[c];
        }
        for (int i = 0; i < CalcOrders; i++) {
            int row = i + 2;
            CalcOrder order = CalcOrderRows[i];
            orders.Cell(row, 1).Value = i + 1;
            orders.Cell(row, 2).Value = order.ProductId;
            orders.Cell(row, 3).Value = order.Quantity;
            orders.Cell(row, 4).Value = order.Region;
            orders.Cell(row, 5).FormulaA1 = $"IFERROR(VLOOKUP(B{row},{productTable},3,FALSE),0)";
            orders.Cell(row, 6).FormulaA1 = $"IFERROR(INDEX({ProductColumn("B")},MATCH(B{row},{ProductColumn("A")},0)),\"{CalcUnknownCategory}\")";
            orders.Cell(row, 7).FormulaA1 = $"C{row}*E{row}";
            orders.Cell(row, 8).FormulaA1 = $"IF(C{row}>={CalcDiscountQuantity},G{row}*IFERROR(VLOOKUP(B{row},{productTable},4,FALSE),0),0)";
            orders.Cell(row, 9).FormulaA1 = $"ROUND(G{row}-H{row},2)";
        }

        var summary = workbook.AddWorksheet(CalcSummarySheet);
        summary.Cell(1, 1).Value = "Category";
        for (int r = 0; r < CalcRegions.Length; r++) {
            summary.Cell(1, 2 + r).Value = CalcRegions[r];
        }
        int countColumn = 2 + CalcRegions.Length;
        summary.Cell(1, countColumn).Value = "Orders";
        for (int i = 0; i < CalcCategories.Length; i++) {
            int row = i + 2;
            summary.Cell(row, 1).Value = CalcCategories[i];
            for (int column = 2; column < countColumn; column++) {
                summary.Cell(row, column).FormulaA1 = $"SUMIFS({OrderColumn("I")},{OrderColumn("F")},$A{row},{OrderColumn("D")},{ColumnName(column)}$1)";
            }
            summary.Cell(row, countColumn).FormulaA1 = $"COUNTIFS({OrderColumn("F")},$A{row})";
        }
        int lastCategoryRow = 1 + CalcCategories.Length;
        int totalsRow = lastCategoryRow + 1;
        summary.Cell(totalsRow, 1).Value = "Total";
        for (int column = 2; column <= countColumn; column++) {
            summary.Cell(totalsRow, column).FormulaA1 = $"SUM({ColumnName(column)}2:{ColumnName(column)}{lastCategoryRow})";
        }
        summary.Cell(CalcGrandTotalRow, 1).Value = "Grand total";
        summary.Cell(CalcGrandTotalRow, CalcStatColumn).FormulaA1 = $"SUM(B2:{ColumnName(countColumn - 1)}{lastCategoryRow})";
        summary.Cell(CalcGrandTotalRow + 1, 1).Value = "Average net";
        summary.Cell(CalcGrandTotalRow + 1, CalcStatColumn).FormulaA1 = $"AVERAGE({OrderColumn("I")})";
        summary.Cell(CalcGrandTotalRow + 2, 1).Value = "Largest net";
        summary.Cell(CalcGrandTotalRow + 2, CalcStatColumn).FormulaA1 = $"MAX({OrderColumn("I")})";
        return Save(workbook);
    }
}
