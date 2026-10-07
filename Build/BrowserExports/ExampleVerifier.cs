using System.Text;
using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;
using OfficeIMO.Excel;

internal static class ExampleVerifier {
    internal static async Task RunAsync(HtmlBrowserSession session, string example, string evidence) {
        IPage page = session.Page;
        var external = new List<string>();
        await session.Context.RouteAsync("**/*", async route => {
            if (route.Request.Url.StartsWith("http", StringComparison.OrdinalIgnoreCase)) {
                external.Add(route.Request.Url);
                await route.AbortAsync();
            } else await route.ContinueAsync();
        });
        foreach (string entry in new[] { "current-view.html", "current-view-bundle.html" }) {
            await page.SetViewportSizeAsync(1280, 800);
            await page.GotoAsync(new Uri(Path.Combine(example, entry)).AbsoluteUri, new() { WaitUntil = WaitUntilState.Load });
            await page.WaitForFunctionAsync("typeof OfficeIMO !== 'undefined' && document.querySelector('table').tBodies[0].rows.length === 3");
            await page.Locator("#filter").FillAsync("Warsaw");
            await page.Locator("#sort").ClickAsync();
            await page.Locator("#move-site").ClickAsync();
            string baseName = Path.GetFileNameWithoutExtension(entry);
            var download = await page.RunAndWaitForDownloadAsync(() => page.Locator("#export-xlsx").ClickAsync());
            string xlsx = Path.Combine(evidence, baseName + ".xlsx");
            await download.SaveAsAsync(xlsx);
            WorkbookVerifier.Verify(xlsx);
            using (var reader = ExcelDocument.OpenDataReader(xlsx, new ExcelReadOptions { HasHeaderRow = false })) {
                WorkbookVerifier.Require(reader.Read() && reader.GetString(0) == "Site" && reader.GetString(1) == "Name", "Example column order differs.");
                WorkbookVerifier.Require(reader.Read() && reader.GetString(0) == "Warsaw" && reader.GetString(1) == "DC03" && Convert.ToDouble(reader.GetValue(2)) == 21, "Example sort/filter differs.");
                WorkbookVerifier.Require(reader.Read() && reader.GetString(1) == "DC01" && !reader.Read(), "Example exported extra rows.");
            }
            await page.ScreenshotAsync(new() { Path = Path.Combine(evidence, baseName + "-wide.png"), FullPage = true });
            await page.Locator("#hide-site").ClickAsync();
            download = await page.RunAndWaitForDownloadAsync(() => page.Locator("#export-csv").ClickAsync());
            string csv = Path.Combine(evidence, baseName + ".csv");
            await download.SaveAsAsync(csv);
            byte[] expected = new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes("Name,Latency (ms)\r\nDC03,21\r\nDC01,12.5\r\n")).ToArray();
            WorkbookVerifier.Require(File.ReadAllBytes(csv).SequenceEqual(expected), "Example visibility/CSV projection differs.");
            await page.SetViewportSizeAsync(390, 740);
            WorkbookVerifier.Require(await page.EvaluateAsync<bool>("document.documentElement.scrollWidth <= innerWidth"), "Compact example scrolls horizontally.");
            await page.ScreenshotAsync(new() { Path = Path.Combine(evidence, baseName + "-compact.png"), FullPage = true });
            await page.Locator("#filter").FillAsync("no matching controller");
            download = await page.RunAndWaitForDownloadAsync(() => page.Locator("#export-xlsx").ClickAsync());
            xlsx = Path.Combine(evidence, baseName + "-empty.xlsx");
            await download.SaveAsAsync(xlsx);
            WorkbookVerifier.Verify(xlsx);
        }
        WorkbookVerifier.Require(external.Count == 0, "Offline example attempted external requests.");
        File.WriteAllText(Path.Combine(evidence, "example-report.json"), JsonSerializer.Serialize(new { passed = true, protocol = "file:", externalRequests = external.Count }));
        Console.WriteLine("Passed inline and bundled file:// example exports.");
    }
}
