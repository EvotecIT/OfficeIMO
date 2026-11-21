using System.Globalization;

/// <summary>The data every library writes, so each lane does the same work everywhere.</summary>
internal static class Workloads {
    // write1m: a header and one million rows of an id, a repeated name, an amount and a flag.
    internal const int LargeRows = 1_000_000;
    internal static readonly string[] LargeHeader = { "Id", "Name", "Amount", "Flag" };
    internal static string LargeName(int i) => Names[i % Names.Length];
    internal static double LargeAmount(int i) => (i % 10_000) * 0.25;
    internal static bool LargeFlag(int i) => i % 3 == 0;
    private static readonly string[] Names = Enumerable.Range(0, 1000).Select(i => "Item " + i.ToString(CultureInfo.InvariantCulture)).ToArray();

    // styles: 20,000 rows by 10 columns, each cell one of 2,000 distinct combinations of fill, font size, bold,
    // italic and border.
    internal const int StyleRows = 20_000;
    internal const int StyleColumns = 10;
    internal const int StyleCount = 2_000;
    internal static int StyleOf(int row, int column) => (row * StyleColumns + column) % StyleCount;
    internal static (byte R, byte G, byte B) Fill(int style) {
        int index = style % 50;
        return ((byte)(40 + index * 4), (byte)(250 - index * 3), (byte)(100 + (index * 37) % 150));
    }
    internal static double FontSize(int style) => 9 + (style / 50) % 4;
    internal static bool Bold(int style) => (style / 200) % 2 == 1;
    internal static bool Italic(int style) => (style / 400) % 2 == 1;
    /// <summary>0 for no border, 1 thin, 2 medium.</summary>
    internal static int Border(int style) => style / 800;

    // merged: 50,000 rows, each with text in A and A:C merged.
    internal const int MergedRows = 50_000;

    // cfdv: 10,000 rows by 3 numeric columns, 2,000 "greater than 50" rules on five-row blocks of B and 2,000
    // whole-number validations on five-row blocks of C.
    internal const int RuleRows = 10_000;
    internal const int RuleBlocks = 2_000;

    // namestables: 200 tables of 50 data rows by 4 columns stacked down the sheet, and 1,000 defined names.
    internal const int Tables = 200;
    internal const int TableRows = 50;
    internal const int DefinedNames = 1_000;
    internal static readonly string[] TableHeader = { "Key", "Value", "Qty", "Note" };
    internal static int TableTop(int table) => table * (TableRows + 2) + 1;

    // recalc: see RecalcFixture.
    internal const int RecalcRows = 20_000;
    internal const int LookupRows = 1_000;
    internal const int Categories = 50;
    internal static double NewLookupValue(int key) => key * 3;

    // autofit: 20,000 rows by 10 columns of text and numbers; only the autofit is timed.
    internal const int AutoFitRows = 20_000;
    internal const int AutoFitColumns = 10;
    internal static bool AutoFitIsText(int column) => column % 2 == 1;
    internal static string AutoFitText(int row, int column) => "Text " + new string('x', (row * column) % 30);
    internal static double AutoFitNumber(int row, int column) => row * column * 1.5;

    // preserve: one value written below the data of the first sheet.
    internal const int PreserveRow = 2_000;
    internal const int PreserveColumn = 26;
    internal const string PreserveText = "edited";

    /// <summary>The Summary!B values once Lookup!B holds <see cref="NewLookupValue"/>.</summary>
    internal static double ExpectedRecalcChecksum() {
        var totals = new double[Categories];
        for (int i = 1; i <= RecalcRows; i++) {
            int key = i % LookupRows + 1;
            double c = NewLookupValue(key) * i;
            double d = c > 1000 ? Math.Round(c / 3, 2, MidpointRounding.AwayFromZero) : c * 2;
            totals[i % Categories] += d;
        }
        return totals.Sum();
    }
}
