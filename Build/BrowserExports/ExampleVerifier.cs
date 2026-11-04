using System.Text;
using System.Text.Json;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlTinkerX;
using Microsoft.Playwright;
using OfficeIMO.Excel;

internal static class ExampleVerifier {
    internal static async Task RunAsync(HtmlBrowserSession session, string example, string evidence, bool translatedLabels = false) {
        IPage page = session.Page;
        var external = new List<string>();
        var errors = new List<string>();
        var consoleErrors = new List<string>();
        EventHandler<string> pageError = (_, message) => errors.Add(message);
        EventHandler<IConsoleMessage> consoleError = (_, message) => { if (message.Type == "error") consoleErrors.Add(message.Text); };
        page.PageError += pageError; page.Console += consoleError;
        var headers = translatedLabels ? new Dictionary<string, string> { ["name"] = "Nazwa", ["site"] = "Lokalizacja", ["latency"] = "Opóźnienie (ms)" }
            : new Dictionary<string, string> { ["name"] = "Name", ["site"] = "Site", ["latency"] = "Latency (ms)" };
        Func<IRoute, Task> blockExternal = async route => {
            if (route.Request.Url.StartsWith("http", StringComparison.OrdinalIgnoreCase)) {
                external.Add(route.Request.Url);
                await route.AbortAsync();
            } else await route.ContinueAsync();
        };
        await session.Context.RouteAsync("**/*", blockExternal);
        try {
        foreach (string entry in new[] { "current-view.html", "current-view-bundle.html" }) {
            await page.SetViewportSizeAsync(1280, 800);
            await page.GotoAsync(new Uri(Path.Combine(example, entry)).AbsoluteUri, new() { WaitUntil = WaitUntilState.Load });
            await page.WaitForFunctionAsync("typeof OfficeIMO !== 'undefined' && document.querySelector('table').tBodies[0].rows.length === 3");
            if (translatedLabels) await page.EvaluateAsync("labels => { for (const cell of document.querySelector('table').tHead.rows[0].cells) { const field = cell.dataset.field; if (!Object.hasOwn(labels, field)) throw Error('Example field identity is missing.'); cell.textContent = labels[field]; } }", headers);
            await page.Locator("#filter").FillAsync("Warsaw");
            await page.Locator("#sort").ClickAsync();
            await page.Locator("#move-site").ClickAsync();
            string baseName = Path.GetFileNameWithoutExtension(entry);
            var download = await page.RunAndWaitForDownloadAsync(() => page.Locator("#export-xlsx").ClickAsync());
            string xlsx = Path.Combine(evidence, baseName + ".xlsx");
            await download.SaveAsAsync(xlsx);
            WorkbookVerifier.Verify(xlsx);
            using (var reader = ExcelDocument.OpenReader(xlsx, new ExcelReaderOptions { HasHeaderRow = false })) {
                WorkbookVerifier.Require(reader.Read() && reader.GetString(0) == headers["site"] && reader.GetString(1) == headers["name"] && reader.GetString(2) == headers["latency"], "Example column order/translated headings differ.");
                WorkbookVerifier.Require(reader.Read() && reader.GetString(0) == "Warsaw" && reader.GetString(1) == "DC03" && Convert.ToDouble(reader.GetValue(2)) == 21, "Example sort/filter differs.");
                WorkbookVerifier.Require(reader.Read() && reader.GetString(1) == "DC01" && !reader.Read(), "Example exported extra rows.");
            }
            VerifyLatencyStyle(xlsx);
            await page.ScreenshotAsync(new() { Path = Path.Combine(evidence, baseName + "-wide.png"), FullPage = true });
            await page.Locator("#hide-site").ClickAsync();
            download = await page.RunAndWaitForDownloadAsync(() => page.Locator("#export-csv").ClickAsync());
            string csv = Path.Combine(evidence, baseName + ".csv");
            await download.SaveAsAsync(csv);
            byte[] expected = Csv(headers["name"] + "," + headers["latency"] + "\r\nDC03,21\r\nDC01,12.5\r\n");
            WorkbookVerifier.Require(File.ReadAllBytes(csv).SequenceEqual(expected), "Example visibility/CSV projection differs.");
            await page.SetViewportSizeAsync(390, 740);
            WorkbookVerifier.Require(await page.EvaluateAsync<bool>("document.documentElement.scrollWidth <= innerWidth"), "Compact example scrolls horizontally.");
            await page.ScreenshotAsync(new() { Path = Path.Combine(evidence, baseName + "-compact.png"), FullPage = true });
            await page.Locator("#filter").FillAsync("no matching controller");
            download = await page.RunAndWaitForDownloadAsync(() => page.Locator("#export-xlsx").ClickAsync());
            xlsx = Path.Combine(evidence, baseName + "-empty.xlsx");
            await download.SaveAsAsync(xlsx);
            WorkbookVerifier.Verify(xlsx);
            using (var reader = ExcelDocument.OpenReader(xlsx, new ExcelReaderOptions { HasHeaderRow = false })) {
                WorkbookVerifier.Require(reader.Read() && reader.GetString(0) == headers["name"] && reader.GetString(1) == headers["latency"] && !reader.Read(), "Empty example must retain only the current visible headings.");
            }
            download = await page.RunAndWaitForDownloadAsync(() => page.Locator("#export-csv").ClickAsync());
            csv = Path.Combine(evidence, baseName + "-empty.csv"); await download.SaveAsAsync(csv);
            WorkbookVerifier.Require(File.ReadAllBytes(csv).SequenceEqual(Csv(headers["name"] + "," + headers["latency"] + "\r\n")), "Empty CSV must retain only the current visible headings.");
        }
        WorkbookVerifier.Require(external.Count == 0, "Offline example attempted external requests.");
        WorkbookVerifier.Require(errors.Count == 0 && consoleErrors.Count == 0, "Offline example runtime/console errors: " + string.Join("; ", errors.Concat(consoleErrors)));
        File.WriteAllText(Path.Combine(evidence, "example-report.json"), JsonSerializer.Serialize(new {
            passed = true, protocol = "file:", labels = translatedLabels ? "explicit Polish test labels selected by stable field" : "default English labels",
            headers, externalRequests = external.Count, errors, consoleErrors, scenarios = new[] { "filter Warsaw", "reverse name", "move Site", "hide Site", "typed and styled latency", "empty visible headings only" },
            captures = new[] { "1280x800 after XLSX", "390x740 after Site hidden" }
        }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine("Passed inline and bundled file:// example exports with " + (translatedLabels ? "translated" : "default") + " headings.");
        } finally {
            page.PageError -= pageError; page.Console -= consoleError;
            await session.Context.UnrouteAsync("**/*", blockExternal);
        }
    }

    private static byte[] Csv(string text) => new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes(text)).ToArray();

    private static void VerifyLatencyStyle(string path) {
        using var package = SpreadsheetDocument.Open(path, false);
        WorkbookPart book = package.WorkbookPart!;
        Cell high = book.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "C2");
        Cell ordinary = book.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "C3");
        WorkbookVerifier.Require((high.DataType is null || high.DataType.Value == CellValues.Number) &&
            double.TryParse(high.CellValue?.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double highValue) && highValue == 21 &&
            (ordinary.DataType is null || ordinary.DataType.Value == CellValues.Number) &&
            double.TryParse(ordinary.CellValue?.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double ordinaryValue) && ordinaryValue == 12.5, "Renamed latency must stay numeric.");
        Stylesheet styles = book.WorkbookStylesPart!.Stylesheet!;
        CellFormat format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)high.StyleIndex!.Value);
        Fill fill = styles.Fills!.Elements<Fill>().ElementAt((int)format.FillId!.Value);
        Font font = styles.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value);
        WorkbookVerifier.Require(fill.Descendants<ForegroundColor>().Single().Rgb?.Value == "FFFCE4D6" && font.Color?.Rgb?.Value == "FF9C0006" && font.Bold is not null && font.Bold.Val?.Value != false,
            "Renamed latency lost its high-value highlight.");
        WorkbookVerifier.Require(styles.NumberingFormats!.Elements<NumberingFormat>().Any(number => number.NumberFormatId!.Value == format.NumberFormatId!.Value && number.FormatCode!.Value == "0.00"),
            "Renamed latency lost its numeric format.");
    }
}
