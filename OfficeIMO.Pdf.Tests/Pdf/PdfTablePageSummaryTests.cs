using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTablePageSummaryTests {
    private static readonly Regex RowMarker = new(@"R(\d{3})", RegexOptions.CultureInvariant);

    private static PdfCreateOptions SmallPage(double height = 320) => new() {
        PageSetup = {
            Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(height)),
            Margins = {
                Left = global::OfficeIMO.OfficeLength.Points(30), Right = global::OfficeIMO.OfficeLength.Points(30),
                Top = global::OfficeIMO.OfficeLength.Points(30), Bottom = global::OfficeIMO.OfficeLength.Points(30)
            }
        },
        DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
    };

    private static decimal Amount(int row) => (row % 7 + 1) * 10.25m + row;

    private static string Format(decimal value) => value.ToString("N2", CultureInfo.InvariantCulture);

    private static List<string[]> InvoiceRows(int count, Func<int, string>? description = null) {
        var rows = new List<string[]> { new[] { "Item", "Description", "Amount" } };
        for (int row = 0; row < count; row++)
            rows.Add(new[] { "R" + row.ToString("000", CultureInfo.InvariantCulture), description?.Invoke(row) ?? "Service", Format(Amount(row)) });
        rows.Add(new[] { "Total", "LAST[{page:amount}]", "TOTAL[{running:amount}]" });
        return rows;
    }

    private static PdfTableStyle InvoiceStyle() {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.FooterRowCount = 1;
        var summary = new PdfTablePageSummary();
        summary.Sum("amount", column: 2);
        summary.AddPageFooterRow("Carried forward", "PS[{page:amount}]", "CF[{running:amount}]");
        summary.AddContinuationHeaderRow("Brought forward", string.Empty, "BF[{broughtForward:amount}]");
        style.PageSummary = summary;
        return style;
    }

    private static List<string> PageTexts(byte[] bytes) {
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        return Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text).ToList();
    }

    private static string? Marker(string text, string name) {
        Match match = Regex.Match(text, Regex.Escape(name) + @"\[([^\]]*)\]");
        return match.Success ? match.Groups[1].Value : null;
    }

    private static int[] RowsOn(string text) =>
        RowMarker.Matches(text).Cast<Match>().Select(match => int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture)).ToArray();

    private static void AssertInvoicePages(List<string> pages, int rowCount) {
        Assert.True(pages.Count > 1, "The table should split across pages.");
        var seen = new List<int>();
        decimal before = 0m;
        for (int page = 0; page < pages.Count; page++) {
            string text = pages[page];
            // A page can hold only the continuation of a split row; that row counts on the page where it starts.
            int[] rows = RowsOn(text);
            seen.AddRange(rows);
            decimal pageSum = rows.Sum(Amount);
            bool last = page == pages.Count - 1;
            if (page == 0) Assert.Null(Marker(text, "BF"));
            else Assert.Equal(Format(before), Marker(text, "BF"));
            if (last) {
                Assert.Null(Marker(text, "CF"));
                Assert.Null(Marker(text, "PS"));
                Assert.Equal(Format(pageSum), Marker(text, "LAST"));
                Assert.Equal(Format(before + pageSum), Marker(text, "TOTAL"));
            } else {
                Assert.Equal(Format(pageSum), Marker(text, "PS"));
                Assert.Equal(Format(before + pageSum), Marker(text, "CF"));
                Assert.Null(Marker(text, "TOTAL"));
            }
            before += pageSum;
        }
        Assert.Equal(Enumerable.Range(0, rowCount), seen);
    }

    [Fact]
    public void PageSubtotalsAndCarriedForwardMatchRowsOnEachPage() {
        byte[] bytes = PdfDocument.Create(SmallPage())
            .Paragraph(paragraph => paragraph.AddText("Invoice lines"))
            .Table(InvoiceRows(60), style: InvoiceStyle())
            .ToBytes();

        List<string> pages = PageTexts(bytes);
        Assert.True(pages.Count >= 3);
        AssertInvoicePages(pages, 60);
    }

    [Fact]
    public void RowsTallerThanTheRemainingSpaceCountOnThePageWhereTheyStart() {
        string Description(int row) => row % 5 == 3
            ? string.Join(" ", Enumerable.Repeat("long wrapped description text", 14))
            : "Service";
        byte[] bytes = PdfDocument.Create(SmallPage())
            .Table(InvoiceRows(40, Description), style: InvoiceStyle())
            .ToBytes();

        List<string> pages = PageTexts(bytes);
        Assert.Contains(pages, page => RowsOn(page).Length == 0);
        AssertInvoicePages(pages, 40);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        // Summary rows take reserved space; nothing is drawn into the bottom margin.
        Assert.All(pdf.GetPages(), page => Assert.All(page.Letters, letter => Assert.True(letter.GlyphRectangle.Bottom >= 29.5,
            $"Text '{letter.Value}' on page {page.Number} sits below the bottom margin.")));
    }

    [Fact]
    public void SinglePageTableShowsOnlyTheClosingFooter() {
        byte[] bytes = PdfDocument.Create(SmallPage(600))
            .Table(InvoiceRows(5), style: InvoiceStyle())
            .ToBytes();

        List<string> pages = PageTexts(bytes);
        Assert.Single(pages);
        decimal total = Enumerable.Range(0, 5).Sum(Amount);
        Assert.Equal(Format(total), Marker(pages[0], "TOTAL"));
        Assert.Equal(Format(total), Marker(pages[0], "LAST"));
        Assert.Null(Marker(pages[0], "CF"));
        Assert.Null(Marker(pages[0], "BF"));
    }

    [Fact]
    public void EmptyTableShowsZeroTotals() {
        byte[] bytes = PdfDocument.Create(SmallPage())
            .Table(InvoiceRows(0), style: InvoiceStyle())
            .ToBytes();

        List<string> pages = PageTexts(bytes);
        Assert.Single(pages);
        Assert.Equal("0.00", Marker(pages[0], "TOTAL"));
        Assert.Equal("0.00", Marker(pages[0], "LAST"));
    }

    [Fact]
    public void RepeatFooterRowsCloseEveryPageFragmentWithItsPageValues() {
        var rows = InvoiceRows(50);
        rows[rows.Count - 1] = new[] { "Page total", "N[{page:count}]", "PS[{page:amount}]" };
        var style = PdfTableStyles.Minimal();
        style.FooterRowCount = 1;
        style.RepeatFooterRows = true;
        style.PageSummary = new PdfTablePageSummary();
        style.PageSummary.Sum("amount", 2);
        style.PageSummary.Add(new PdfTableAggregate("count", 0, PdfTableAggregateFunction.Count));

        List<string> pages = PageTexts(PdfDocument.Create(SmallPage()).Table(rows, style: style).ToBytes());

        Assert.True(pages.Count > 1);
        var seen = new List<int>();
        foreach (string page in pages) {
            int[] onPage = RowsOn(page);
            seen.AddRange(onPage);
            Assert.Equal(Format(onPage.Sum(Amount)), Marker(page, "PS"));
            Assert.Equal(onPage.Length.ToString(CultureInfo.InvariantCulture), Marker(page, "N"));
            Assert.Single(Regex.Matches(page, "Page total").Cast<Match>());
        }
        Assert.Equal(Enumerable.Range(0, 50), seen);
    }

    [Fact]
    public void RepeatFooterRowsWithoutAggregatesRepeatStaticFooters() {
        var rows = new List<string[]> { new[] { "Item", "Value" } };
        rows.AddRange(Enumerable.Range(0, 50).Select(row => new[] { "R" + row.ToString("000", CultureInfo.InvariantCulture), "x" }));
        rows.Add(new[] { "Continued table footer", "end" });
        var style = PdfTableStyles.Minimal();
        style.FooterRowCount = 1;
        style.RepeatFooterRows = true;

        List<string> pages = PageTexts(PdfDocument.Create(SmallPage()).Table(rows, style: style).ToBytes());

        Assert.True(pages.Count > 1);
        Assert.All(pages, page => Assert.Contains("Continued table footer", page));
        Assert.Equal(Enumerable.Range(0, 50), pages.SelectMany(RowsOn));
    }

    [Fact]
    public void MinimumMaximumAndCultureFormattingUseTheCallerSettings() {
        var culture = CultureInfo.GetCultureInfo("de-DE");
        var rows = new List<string[]> { new[] { "Item", "Amount" } };
        rows.AddRange(Enumerable.Range(0, 40).Select(row => new[] {
            "R" + row.ToString("000", CultureInfo.InvariantCulture), (Amount(row) * 100m).ToString("N2", culture) }));
        rows.Add(new[] { "MIN[{running:low}]", "MAX[{running:high}]" });
        var style = PdfTableStyles.Minimal();
        style.FooterRowCount = 1;
        style.PageSummary = new PdfTablePageSummary();
        style.PageSummary.Add(new PdfTableAggregate("low", 1, PdfTableAggregateFunction.Minimum) { FormatProvider = culture });
        style.PageSummary.Add(new PdfTableAggregate("high", 1, PdfTableAggregateFunction.Maximum) { FormatProvider = culture, Format = "N0" });
        style.PageSummary.Sum("amount", 1, "N2", culture);
        style.PageSummary.AddPageFooterRow("Carried forward", "CF[{running:amount}]");

        List<string> pages = PageTexts(PdfDocument.Create(SmallPage()).Table(rows, style: style).ToBytes());

        decimal[] values = Enumerable.Range(0, 40).Select(row => Amount(row) * 100m).ToArray();
        string last = pages[pages.Count - 1];
        Assert.Equal(values.Min().ToString("N2", culture), Marker(last, "MIN"));
        Assert.Equal(values.Max().ToString("N0", culture), Marker(last, "MAX"));
        decimal running = 0m;
        foreach (string page in pages.Take(pages.Count - 1)) {
            running += RowsOn(page).Sum(row => Amount(row) * 100m);
            Assert.Equal(running.ToString("N2", culture), Marker(page, "CF"));
        }
    }

    [Fact]
    public void ValueSelectorReplacesCellTextParsing() {
        var rows = InvoiceRows(30);
        var style = InvoiceStyle();
        style.PageSummary!.Aggregates.Clear();
        style.PageSummary.Add(new PdfTableAggregate("amount", 2) { ValueSelector = bodyRow => bodyRow });

        List<string> pages = PageTexts(PdfDocument.Create(SmallPage()).Table(rows, style: style).ToBytes());

        Assert.Equal(Format(Enumerable.Range(0, 30).Sum()), Marker(pages[pages.Count - 1], "TOTAL"));
    }

    [Fact]
    public void GeneratedBytesAreDeterministic() {
        byte[] Render() => PdfDocument.Create(SmallPage())
            .Table(InvoiceRows(45, row => row % 9 == 4 ? string.Join(" ", Enumerable.Repeat("wrapped", 40)) : "Service"), style: InvoiceStyle())
            .ToBytes();

        Assert.Equal(Render(), Render());
    }

    [Fact]
    public void UndeclaredAggregateTokensAreRejected() {
        var style = InvoiceStyle();
        style.PageSummary!.AddPageFooterRow("x", "{page:missing}", "y");

        Assert.Throws<ArgumentException>(() => PdfDocument.Create(SmallPage()).Table(InvoiceRows(3), style: style).ToBytes());
    }

    [Fact]
    public void DeferredTablesRejectPageSummaries() {
        Assert.Throws<NotSupportedException>(() => PdfDocument.Create(SmallPage())
            .TableDeferred(() => InvoiceRows(3), style: InvoiceStyle())
            .ToBytes());
    }

    [Fact]
    public void InvoiceComponentCarriesLineTotalsAcrossPages() {
        var lines = Enumerable.Range(0, 90).Select(line => new PdfInvoiceLine("R" + line.ToString("000", CultureInfo.InvariantCulture), 1M, line + 1M)).ToArray();
        var invoice = new PdfInvoiceComponent("INV-7", new DateTime(2026, 11, 1), new PdfInvoiceParty("Seller Ltd"),
            new PdfInvoiceParty("Customer Ltd"), lines, "EUR") { CarryTotalsAcrossPages = true };

        List<string> pages = PageTexts(PdfDocument.Create().Component(invoice).ToBytes());

        int lastTablePage = pages.FindIndex(page => page.Contains("R089"));
        decimal running = 0m;
        for (int page = 0; page <= lastTablePage; page++) {
            if (page > 0) {
                Assert.Contains("Brought forward", pages[page]);
                Assert.Contains(Format(running) + " EUR", pages[page]);
            }
            running += RowsOn(pages[page]).Sum(line => line + 1m);
            if (page < lastTablePage) {
                Assert.Contains("Carried forward", pages[page]);
                Assert.Contains(Format(running) + " EUR", pages[page]);
            }
        }
        Assert.True(lastTablePage > 0);
        Assert.Equal(Enumerable.Range(1, 90).Sum(), running);
    }

    [Fact]
    public void StyleCloneCopiesThePageSummary() {
        PdfTableStyle style = InvoiceStyle();
        style.RepeatFooterRows = true;
        PdfTableStyle clone = style.Clone();

        Assert.True(clone.RepeatFooterRows);
        Assert.NotSame(style.PageSummary, clone.PageSummary);
        Assert.Equal("amount", Assert.Single(clone.PageSummary!.Aggregates).Name);
        Assert.Single(clone.PageSummary.PageFooterRows);
        Assert.Single(clone.PageSummary.ContinuationHeaderRows);
    }
}
