using System.Text.RegularExpressions;
using Microsoft.Playwright;

internal static class ChartProof {
    private const string NativeFigure = "figure[data-officeimo-chart]";
    private const string Title = "Station temperature heatmap";

    public static async Task VerifyHeatmapAsync(IPage page, bool native, string? preview) {
        var heading = page.GetByRole(AriaRole.Heading, new() { NameRegex = new Regex("^Plotly evidence HEATMAP") });
        await heading.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 8000 });
        if (native) {
            await VerifyOfficeHeatmapAsync(page, GetOfficeFigure(page, Title), preview);
        } else {
            await VerifyPlotlyHeatmapAsync(page, preview);
        }
        Console.WriteLine("Verified saved heatmap x=2, y=25, z=18.5, zoom and reset.");
    }

    public static async Task VerifyNotebookAsync(IPage page, string? widePreview, string? compactPreview) {
        await VerifyNotebookLayoutAsync(page, widePreview);
        var mesh = GetOfficeFigure(page, "Surveyed ridge mesh").Locator("svg");
        var width = await mesh.EvaluateAsync<double[]>("svg => [svg.getBoundingClientRect().width, svg.viewBox.baseVal.width]");
        Require(Math.Abs(width[0] - 700) < 1 && Math.Abs(width[1] * 4 / 3 - 700) < 1,
            $"The saved 700px mesh should occupy 700 CSS pixels at the wide viewport. Found {width[0]}px and a {width[1]}pt viewBox.");
        await page.SetViewportSizeAsync(390, 844);
        await VerifyNotebookLayoutAsync(page, compactPreview);
        Console.WriteLine("Verified five offline chart viewers at 1080px and 390px, with keyboard and pointer navigation, exact heatmap values, and no horizontal overflow.");
    }

    private static async Task VerifyNotebookLayoutAsync(IPage page, string? preview) {
        await VerifyOfficeHeatmapAsync(page, GetOfficeFigure(page, Title), preview);
        foreach (var title in new[] { "Surveyed ridge mesh", "Hudson survey route", "Hudson sampling regions", "Field survey comparison" }) {
            var figure = GetOfficeFigure(page, title);
            await VerifyOfficeNavigationAsync(figure);
        }
        Require(await page.Locator(NativeFigure).CountAsync() == 5, "Expected five independent saved chart viewers.");
        var fits = await page.EvaluateAsync<bool>("""
            () => {
              const width = document.documentElement.clientWidth;
              const items = document.querySelectorAll("figure[data-officeimo-chart], figure[data-officeimo-chart] svg, [data-officeimo-chart-tools]");
              return document.documentElement.scrollWidth <= width + 1 && Array.from(items).every(item => {
                const bounds = item.getBoundingClientRect();
                return bounds.left >= -1 && bounds.right <= width + 1 && item.scrollWidth <= item.clientWidth + 1;
              });
            }
            """);
        Require(fits, "A chart or navigation toolbar overflows the viewport.");
    }

    private static ILocator GetOfficeFigure(IPage page, string title) => page.Locator(NativeFigure)
        .Filter(new() { Has = page.GetByRole(AriaRole.Img, new() { Name = title, Exact = true }) });

    private static async Task VerifyOfficeHeatmapAsync(IPage page, ILocator figure, string? preview) {
        var svg = figure.Locator("svg");
        await svg.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 10000 });
        await svg.ScrollIntoViewIfNeededAsync();
        var target = await figure.EvaluateAsync<float[]>("""
            figure => {
              const svg = figure.querySelector("svg");
              const marks = JSON.parse(figure.querySelector("template[data-officeimo-chart-marks]").content.textContent);
              const [x, y, width, height] = marks.find(mark => mark[4] === "x: 2 · y: 25 · z: 18.5");
              const point = new DOMPoint(x + width / 2, y + height / 2).matrixTransform(svg.getScreenCTM());
              return [point.x, point.y];
            }
            """);
        await page.Mouse.MoveAsync(target[0], target[1]);
        var output = figure.Locator("output[data-officeimo-chart-value]");
        await output.GetByText("x: 2 · y: 25 · z: 18.5", new() { Exact = true })
            .WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 5000 });
        await CaptureAsync(page, preview);

        await figure.GetByRole(AriaRole.Button, new() { Name = "Next data point", Exact = true }).ClickAsync(new() { Timeout = 5000 });
        var next = await output.TextContentAsync();
        Require(!string.IsNullOrEmpty(next) && next != "x: 2 · y: 25 · z: 18.5", "Next data point did not change the inspected value.");
        await VerifyOfficeNavigationAsync(figure, page);
    }

    private static async Task VerifyOfficeNavigationAsync(ILocator figure, IPage? page = null) {
        var svg = figure.Locator("svg");
        await svg.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 10000 });
        var initial = await svg.GetAttributeAsync("viewBox") ?? throw new InvalidOperationException("Office plot requires a viewBox.");
        await figure.GetByRole(AriaRole.Button, new() { Name = "Zoom in", Exact = true }).ClickAsync(new() { Timeout = 5000 });
        var zoomed = await svg.GetAttributeAsync("viewBox");
        Require(zoomed is not null && zoomed != initial, "Office heatmap did not zoom.");
        await svg.PressAsync("ArrowRight", new() { Timeout = 5000 });
        var panned = await svg.GetAttributeAsync("viewBox");
        Require(panned != zoomed, "Office chart did not pan by keyboard.");
        if (page is not null) {
            await svg.ScrollIntoViewIfNeededAsync();
            var bounds = await svg.BoundingBoxAsync() ?? throw new InvalidOperationException("Chart bounds missing.");
            await page.Mouse.MoveAsync(bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2);
            await page.Mouse.DownAsync();
            await page.Mouse.MoveAsync(bounds.X + bounds.Width / 2 + 35, bounds.Y + bounds.Height / 2, new() { Steps = 3 });
            await page.Mouse.UpAsync();
            Require(await svg.GetAttributeAsync("viewBox") != panned, "Office chart did not pan by dragging.");
        }
        await figure.GetByRole(AriaRole.Button, new() { Name = "Reset view", Exact = true }).ClickAsync(new() { Timeout = 5000 });
        Require(await svg.GetAttributeAsync("viewBox") == initial, "Office heatmap did not restore its saved viewport.");
    }

    private static async Task VerifyPlotlyHeatmapAsync(IPage page, string? preview) {
        var plot = page.Locator(".plotly-graph-div");
        var title = plot.GetByText(Title, new() { Exact = true });
        await title.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 10000 });
        var image = plot.Locator(".heatmaplayer image");
        var box = await image.BoundingBoxAsync() ?? throw new InvalidOperationException("Expected a drawn heatmap.");
        await image.ScrollIntoViewIfNeededAsync();
        var region = await image.BoundingBoxAsync() ?? box;
        await page.Mouse.MoveAsync(region.X + 0.375f * region.Width, region.Y + 0.5f * region.Height);
        var hover = plot.Locator(".hoverlayer .hovertext");
        await hover.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 5000 });
        ValidateValues(await hover.TextContentAsync());
        await CaptureAsync(page, preview);

        var initial = await ReadPlotlyRangeAsync(plot);
        await ClickPlotlyControlAsync(plot, "Zoom in");
        Require(await ReadPlotlyRangeAsync(plot) != initial, "Plotly heatmap did not zoom.");
        await ClickPlotlyControlAsync(plot, "Reset axes");
        Require(await ReadPlotlyRangeAsync(plot) == initial, "Plotly heatmap did not restore its saved x-axis range.");
    }

    private static Task<string> ReadPlotlyRangeAsync(ILocator plot) =>
        plot.EvaluateAsync<string>("gd => JSON.stringify(gd._fullLayout.xaxis.range)");

    private static Task ClickPlotlyControlAsync(ILocator plot, string title) =>
        plot.Locator($".modebar-btn[data-title='{title}']").ClickAsync(new() { Timeout = 5000 });

    private static void ValidateValues(string? text) {
        string values = text ?? string.Empty;
        Require(new[] { "2", "25", "18.5" }.All(value => Regex.IsMatch(values, $@"(?<![\d.]){Regex.Escape(value)}(?![\d.])")),
            $"Expected heatmap hover values x=2, y=25, z=18.5, found `{text}`.");
    }

    private static Task CaptureAsync(IPage page, string? preview) => preview is null
        ? Task.CompletedTask
        : page.ScreenshotAsync(new() { Path = preview, FullPage = true });

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
