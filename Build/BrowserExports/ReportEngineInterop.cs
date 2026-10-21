using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;

/// <summary>Actual report-workbench UI qualification through HtmlTinkerX-owned browser sessions.</summary>
internal static class ReportEngineInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        var uri = new Uri(args.FirstOrDefault(a => a.StartsWith("--site-url=", StringComparison.Ordinal))?[11..]
            ?? throw new ArgumentException("--report-engine requires --site-url=<loopback report-engine URL>."));
        if (!uri.IsLoopback || uri.Scheme != "http") throw new ArgumentException("Workbench qualification requires loopback HTTP.");
        using var http = new HttpClient();
        string[] paths = [uri.AbsolutePath, "/js/report-engine-job.mjs", "/js/analytical-report-job.mjs", "/js/report-engine.mjs", "/js/report-engine-loader.mjs", "/js/report-engine-worker.mjs", "/assets/officeimo/officeimo.mjs", "/assets/officeimo/officeimo-execution.mjs", "/css/report-engine.css"];
        async Task<Dictionary<string, string>> ServedHashesAsync() {
            var hashes = new Dictionary<string, string>();
            foreach (string path in paths) hashes.Add(path, Convert.ToHexString(SHA256.HashData(await http.GetByteArrayAsync(new Uri(uri, path)))).ToLowerInvariant());
            return hashes;
        }
        var served = await ServedHashesAsync();
        bool visualOnly = args.Contains("--report-engine-visual-only");
        bool lifecycleOnly = args.Contains("--report-engine-lifecycle-only");
        bool refreshOnly = args.Contains("--report-engine-refresh-only");
        bool hierarchicalOnly = args.Contains("--report-engine-hierarchical-only");
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            await using var session = await HtmlBrowser.OpenSessionAsync(uri.AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timezone = "Europe/Warsaw", Timeout = 120000 });
            var page = session.Page;
            var errors = new List<string>(); var console = new List<object>(); var network = new List<NetworkFailure>();
            page.PageError += (_, error) => errors.Add(error);
            page.Console += (_, message) => { if (message.Type is "error" or "warning") console.Add(new { type = message.Type, text = message.Text }); };
            page.Response += (_, response) => { if (response.Status >= 400) network.Add(new(response.Url, response.Status)); };
            var files = new List<object>(); var states = new List<object>();
            if (lifecycleOnly) {
                object lifecycle = await VerifyLifecycleAsync(page, output, args.Contains("--expect-lifecycle-defect"));
                var lifecycleReport = new { engine = engine.ToString(), browser = session.Browser?.Version, site = uri.AbsoluteUri, mode = "lifecycle-only", served, lifecycle, pageErrors = errors, console, network };
                Require(errors.Count == 0, "Lifecycle check had uncaught page errors: " + string.Join("; ", errors));
                reports.Add(lifecycleReport); File.WriteAllText(Path.Combine(output, "lifecycle.json"), JsonSerializer.Serialize(lifecycleReport, new JsonSerializerOptions { WriteIndented = true }));
                Console.WriteLine($"Report workbench lifecycle qualified {engine}; actual navigation and explicit persisted-event Blob boundary recorded.");
                continue;
            }
            await page.ReloadAsync(new() { WaitUntil = WaitUntilState.Load });
            await ReadyAsync(page);
            Require(await page.Locator("#report-preview tr").CountAsync() == 8, "Default preview must contain eight rows.");
            Require(await page.Locator("#report-preview tr").Nth(1).Locator("td").Nth(3).InnerTextAsync() == "4.9 ms", "Preview numeric rounding differs from PDF.");
            Require(await page.EvaluateAsync<bool>("document.querySelector('#report-status').getAttribute('role') === 'status' && document.querySelector('#report-status').getAttribute('aria-live') === 'polite'"), "Status is not a live region.");

            if (hierarchicalOnly) {
                await page.AddScriptTagAsync(new() { Path = Path.Combine(repository, "Build", "BrowserExports", "hierarchical-report-contracts.js") });
                await VerifyHierarchicalAsync(page, engine, output, files, states);
                Require(errors.Count == 0 && network.Count == 0, "Hierarchical workbench had page or HTTP failures.");
                var hierarchical = new { engine = engine.ToString(), browser = session.Browser?.Version, mode = "hierarchical-only", site = uri.AbsoluteUri, served, files, states, pageErrors = errors, console, network };
                reports.Add(hierarchical); File.WriteAllText(Path.Combine(output, "hierarchical.json"), JsonSerializer.Serialize(hierarchical, new JsonSerializerOptions { WriteIndented = true }));
                Console.WriteLine($"Hierarchical workbench passed {engine}: localized/empty XLSX and PDF downloads, cancellation, CSV recovery and responsive controls.");
                continue;
            }

            if (!visualOnly) {
            foreach (int count in refreshOnly ? new[] { 100 } : new[] { 100, 0 }) foreach (string format in refreshOnly ? new[] { "xlsx", "pdf" } : new[] { "xlsx", "pdf", "csv" }) {
                await SelectAsync(page, count, format);
                await GenerateAsync(page);
                Require(await page.Locator("#report-download").IsVisibleAsync(), "Successful report has no download.");
                files.Add(await DownloadAsync(page, output, "report-" + count + "." + format, count));
                if (count == 0) Require(await page.EvaluateAsync<bool>("document.querySelector('#report-preview').rows.length === 1 && document.querySelector('#report-preview td').colSpan === 4"), "Empty preview lost its four-column placeholder.");
            }

            await SelectAsync(page, 50000, "pdf");
            string oldDownload = await page.Locator("#report-download").GetAttributeAsync("href") ?? "";
            await page.Locator("#report-generate").ClickAsync();
            await page.WaitForFunctionAsync("document.querySelector('#report-progress').value > 0 && document.querySelector('#report-form').getAttribute('aria-busy') === 'true'", null, new() { Timeout = 120000 });
            double processed = await page.Locator("#report-progress").EvaluateAsync<double>("element => element.value");
            await page.Locator("#report-cancel").ClickAsync(); await ReadyAsync(page);
            Require(await page.Locator("#report-progress").EvaluateAsync<double>("element => element.value") == 0 && await page.Locator("#report-download").GetAttributeAsync("href") == oldDownload, "Cancelled work replaced the previous completed download.");
            states.Add(new { cancellation = "after observable DOM row progress", processed, partialDownloadCreated = false, recovered = true });
            await SelectAsync(page, 100, "csv"); await GenerateAsync(page);
            files.Add(await DownloadAsync(page, output, "recovery-after-cancel.csv", 100));
            if (refreshOnly) {
                Require(errors.Count == 0 && network.Count == 0, "Final artifact/cancellation refresh had page or HTTP failures.");
                var refresh = new { engine = engine.ToString(), browser = session.Browser?.Version, mode = "refresh-only", site = uri.AbsoluteUri, served, files, states, pageErrors = errors, console, network };
                reports.Add(refresh); File.WriteAllText(Path.Combine(output, "refresh.json"), JsonSerializer.Serialize(refresh, new JsonSerializerOptions { WriteIndented = true }));
                Console.WriteLine($"Report workbench refresh passed {engine}: current XLSX/PDF UI downloads, cancellation after DOM progress, and CSV recovery.");
                continue;
            }

            const string fontRoute = "**/assets/report-fonts/Carlito-Regular.ttf";
            int injected = 0;
            await HtmlBrowser.RegisterRouteAsync(session, fontRoute, route => { injected++; return route.FulfillAsync(new() { Status = 503, ContentType = "text/plain", Body = "Deliberate workbench font-fetch boundary failure" }); });
            await SelectAsync(page, 100, "pdf");
            oldDownload = await page.Locator("#report-download").GetAttributeAsync("href") ?? "";
            await page.Locator("#report-generate").ClickAsync(); await ReadyAsync(page);
            Require(injected == 1 && await page.Locator("#report-progress").EvaluateAsync<double>("element => element.value") == 0 && await page.Locator("#report-download").GetAttributeAsync("href") == oldDownload, "Font fetch failure did not preserve the recoverable form/download state.");
            states.Add(new { fontHttpFailure = 503, requests = injected, status = await page.Locator("#report-status").InnerTextAsync() });
            await HtmlBrowser.UnregisterRouteAsync(session, fontRoute);
            await GenerateAsync(page); files.Add(await DownloadAsync(page, output, "recovery-after-font.pdf", 100));

            const string moduleRoute = "**/js/report-engine.mjs";
            int moduleFailures = 0;
            await HtmlBrowser.RegisterRouteAsync(session, moduleRoute, route => { moduleFailures++; return route.FulfillAsync(new() { Status = 503, ContentType = "text/javascript", Body = "" }); });
            await page.ReloadAsync();
            await page.WaitForFunctionAsync("!document.querySelector('#report-generate').disabled && document.querySelector('#report-generate').type === 'button'");
            Require(moduleFailures == 1 && await page.Locator("#report-cancel").IsDisabledAsync() && !await page.Locator("#report-download").IsVisibleAsync(), "Module bootstrap failure exposed unusable export controls.");
            await HtmlBrowser.UnregisterRouteAsync(session, moduleRoute);
            await page.Locator("#report-generate").ClickAsync(); await ReadyAsync(page);
            Require(await page.Locator("#report-generate").GetAttributeAsync("type") == "submit" && await page.Locator("#report-preview tr").CountAsync() == 8, "Reload controls did not restore the actual report form.");
            states.Add(new { moduleHttpFailure = 503, requests = moduleFailures, reloadControlsRecovered = true });

            await page.Locator("#report-rows").FocusAsync(); await page.Keyboard.PressAsync("Tab");
            Require(await page.EvaluateAsync<string>("document.activeElement.id") == "report-format", "Keyboard order skips format selection.");
            await page.Keyboard.PressAsync("Tab");
            string tabTarget = await page.EvaluateAsync<string>("document.activeElement.id");
            if (tabTarget != "report-generate" && engine == HtmlBrowserEngine.WebKit) {
                // Safari's default macOS Tab policy omits buttons; Option+Tab includes them without changing user preferences.
                await page.Locator("#report-format").FocusAsync(); await page.Keyboard.PressAsync("Alt+Tab");
            }
            Require(await page.EvaluateAsync<string>("document.activeElement.id") == "report-generate", "Keyboard order skips Generate: " + await page.EvaluateAsync<string>("document.activeElement.outerHTML"));
            Require(await page.EvaluateAsync<bool>("parseFloat(getComputedStyle(document.activeElement).outlineWidth) >= 2 && getComputedStyle(document.activeElement).outlineStyle !== 'none'"), "Keyboard focus lacks a visible outline.");
            string keyboardFormat = await page.Locator("#report-format").InputValueAsync();
            string priorKeyboardDownload = await page.Locator("#report-download").GetAttributeAsync("href") ?? "";
            await page.Keyboard.PressAsync("Enter"); await NewDownloadAsync(page, priorKeyboardDownload);
            files.Add(await DownloadAsync(page, output, "keyboard-100." + keyboardFormat, 100));
            states.Add(new { keyboard = "rows → format → generate → Enter", nativeTabTarget = tabTarget, webKitOptionTab = engine == HtmlBrowserEngine.WebKit && tabTarget != "report-generate", visibleFocus = true });
            }

            foreach (string theme in new[] { "light", "dark" }) {
                await page.EvaluateAsync("theme => { localStorage.setItem('imo-theme', theme); localStorage.setItem('imo-bg', 'blueprint'); }", theme);
                await page.ReloadAsync(); await ReadyAsync(page);
                foreach ((int width, int height, string label) in new[] { (1280, 900, "wide"), (390, 844, "phone") }) {
                    await page.SetViewportSizeAsync(width, height); await page.EvaluateAsync("() => document.fonts.ready.then(() => true)");
                    await page.Locator("#report-generate").FocusAsync();
                    if (label == "phone" && theme == "light") {
                        await page.Locator(".imo-report-lab__table-scroll").FocusAsync(); await page.Keyboard.PressAsync("ArrowRight");
                        bool keyboardScroll = true;
                        try { await page.WaitForFunctionAsync("document.querySelector('.imo-report-lab__table-scroll').scrollLeft > 0 || document.querySelector('.imo-report-lab__table-scroll').scrollWidth <= document.querySelector('.imo-report-lab__table-scroll').clientWidth", null, new() { Timeout = 2000 }); }
                        catch (TimeoutException) {
                            Require(engine == HtmlBrowserEngine.WebKit, "Preview keyboard scrolling failed outside the WebKit native policy boundary.");
                            var probe = await session.Context.NewPageAsync();
                            try {
                                await probe.SetContentAsync("<div id='native-scroll' tabindex='0' style='width:100px;overflow:auto;border:1px solid'><div style='width:800px;height:80px'>Native overflow reference</div></div>");
                                await probe.Locator("#native-scroll").FocusAsync(); await probe.Keyboard.PressAsync("ArrowRight");
                                await probe.WaitForTimeoutAsync(300);
                                Require(await probe.Locator("#native-scroll").EvaluateAsync<double>("element => element.scrollLeft") == 0, "The workbench scroll failed while WebKit's independent native overflow reference succeeded.");
                            } finally { await probe.CloseAsync(); }
                            keyboardScroll = false;
                            await page.Locator(".imo-report-lab__table-scroll").HoverAsync(); await page.Mouse.WheelAsync(120, 0);
                            await page.WaitForFunctionAsync("document.querySelector('.imo-report-lab__table-scroll').scrollLeft > 0", null, new() { Timeout = 2000 });
                        }
                        await page.Locator(".imo-report-lab__table-scroll").EvaluateAsync("element => element.scrollTo({ left: 0, behavior: 'instant' })");
                        states.Add(new { previewKeyboardHorizontalScroll = keyboardScroll, webKitNativeReferenceAlsoOmitsArrowScroll = !keyboardScroll, horizontalPointerScroll = !keyboardScroll, bodyRemainedWithinViewport = true });
                    }
                    await ScrollTopAsync(page);
                    JsonElement layout = await LayoutAsync(page);
                    Require(layout.GetProperty("overflow").GetInt32() <= 1, $"{theme}/{label} has body overflow.");
                    Require(layout.GetProperty("controlsVisible").GetBoolean() && layout.GetProperty("theme").GetString() == theme, $"{theme}/{label} controls/theme differ.");
                    Require(layout.GetProperty("selectHeights").EnumerateArray().All(value => value.GetDouble() >= 44), $"{theme}/{label} select controls are below 44px: " + layout.GetProperty("selectHeights").GetRawText());
                    await page.ScreenshotAsync(new() { Path = Path.Combine(output, label + "-" + theme + ".png"), FullPage = true });
                    states.Add(new { theme, width, height, layout });
                }
            }
            await page.Locator("#report-rows").SelectOptionAsync("0");
            await ScrollTopAsync(page);
            await page.ScreenshotAsync(new() { Path = Path.Combine(output, "phone-empty-dark.png"), FullPage = true });
            Require(await page.Locator(".imo-footer a[href='/report-engine/']").CountAsync() > 0, "Workbench footer lacks its report link.");
            var neighbor = new Uri(uri, "/convert/"); await page.GotoAsync(neighbor.AbsoluteUri);
            Require(await page.Locator(".imo-prose a[href='/report-engine/']").CountAsync() > 0, "Converter neighbor lacks its workbench content link.");
            await page.Locator(".imo-prose a[href='/report-engine/']").First.ClickAsync(); await ReadyAsync(page);
            Require(new Uri(page.Url).AbsolutePath == uri.AbsolutePath, "Converter CTA did not reach the workbench.");
            states.Add(new { neighbor = neighbor.AbsoluteUri, converterContentCta = true, workbenchFooterLink = true, ctaFollowed = true, converterTemplateFooter = "absent in existing template; unchanged" });
            Require(errors.Count == 0, "Uncaught page errors: " + string.Join("; ", errors));
            var expectedFailures = new[] { new Uri(uri, "/assets/report-fonts/Carlito-Regular.ttf").AbsoluteUri, new Uri(uri, "/js/report-engine.mjs").AbsoluteUri };
            Require(network.All(failure => failure.Status == 503 && expectedFailures.Contains(failure.Url)), "Unexpected HTTP failure: " + JsonSerializer.Serialize(network));
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, mode = visualOnly ? "visual-only" : "full", site = uri.AbsoluteUri, served, files, states, pageErrors = errors, console, network };
            reports.Add(report); File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine(visualOnly ? $"Report workbench visuals passed {engine}: responsive/theme/focus/CTA checks; artifact proof retained separately."
                : $"Report workbench passed {engine}: {files.Count} UI downloads independently read; cancel/font retry; responsive/theme/focus checks.");
        }
        Require(served.OrderBy(pair => pair.Key).SequenceEqual((await ServedHashesAsync()).OrderBy(pair => pair.Key)), "The served workbench candidate changed during qualification.");
        File.WriteAllText(Path.Combine(evidence, "report-engine-interop.json"), JsonSerializer.Serialize(new {
            exampleSha256 = Hash(Path.Combine(repository, "OfficeIMO.JavaScript.Examples", "report-engine-job.mjs")),
            analyticalExampleSha256 = Hash(Path.Combine(repository, "OfficeIMO.JavaScript.Examples", "analytical-report-job.mjs")), served, reports
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static async Task SelectAsync(IPage page, int count, string format) {
        await page.Locator("#report-rows").SelectOptionAsync(count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        await page.Locator("#report-format").SelectOptionAsync(format);
    }
    private static async Task ReadyAsync(IPage page) => await page.WaitForFunctionAsync("document.querySelector('#report-form').getAttribute('aria-busy') === 'false' && !document.querySelector('#report-generate').disabled", null, new() { Timeout = 120000 });
    private static async Task GenerateAsync(IPage page) {
        string prior = await page.Locator("#report-download").GetAttributeAsync("href") ?? "";
        await page.Locator("#report-generate").ClickAsync(); await NewDownloadAsync(page, prior);
    }
    private static async Task NewDownloadAsync(IPage page, string prior) => await page.WaitForFunctionAsync("prior => { const link=document.querySelector('#report-download'), progress=document.querySelector('#report-progress'); return link.getAttribute('href') !== prior && !link.hidden && progress.value === progress.max && document.querySelector('#report-form').getAttribute('aria-busy') === 'false'; }", prior, new() { Timeout = 120000 });
    private static async Task<object> DownloadAsync(IPage page, string output, string name, int count, string? contract = null) {
        var download = await page.RunAndWaitForDownloadAsync(() => page.Locator("#report-download").ClickAsync(), new() { Timeout = 120000 });
        string path = Path.Combine(output, name); await download.SaveAsAsync(path);
        string? failure = await download.FailureAsync();
        Require(failure is null && download.SuggestedFilename == "service-health" + Path.GetExtension(name), $"Download differs: expected service-health{Path.GetExtension(name)}, actual {download.SuggestedFilename}, failure {failure ?? "none"}.");
        if (contract is null) return ReportEngineVerifier.Verify(path, count);
        File.WriteAllText(path + ".json", contract);
        return HierarchicalReportVerifier.Verify(path);
    }
    private static async Task VerifyHierarchicalAsync(IPage page, HtmlBrowserEngine engine, string output, List<object> files, List<object> states) {
        async Task SelectHierarchyAsync(int count, string format, string language) {
            await SelectAsync(page, count, format);
            await page.Locator("#report-layout").SelectOptionAsync("grouped");
            Require(await page.Locator("#report-language").IsEnabledAsync(), "Site grouping lost language selection.");
            await page.Locator("#report-layout").SelectOptionAsync("hierarchical");
            await page.Locator("#report-language").SelectOptionAsync(language);
            Require(await page.Locator("#report-language").IsEnabledAsync(), "Hierarchical reports lost language selection.");
            if (count > 0) Require(await page.Locator("#report-preview tr").First.Locator("td").Nth(3).InnerTextAsync() == "—", "Null latency was displayed as a numeric zero.");
        }
        async Task<string> ContractAsync(int count, string language) => await page.EvaluateAsync<string>("""
            async ({count,language}) => {
              const { analyticalRows, analyticalReportText } = await import('/js/analytical-report-job.mjs');
              const { reportColumns } = await import('/js/report-engine-job.mjs');
              const text = { ...analyticalReportText[language], scope: language === 'pl'
                ? `Wybrany widok usług · ${count.toLocaleString('pl')} rekordów.` : `Selected service view · ${count.toLocaleString('en')} services.` };
              const columns = reportColumns.map(column => ({...column, header: text.headers[column.key] ?? column.header}));
              return JSON.stringify(hierarchicalExpectedContract([...analyticalRows(count)], columns, ['portfolio','environment','owner'], text, language));
            }
            """, new { count, language });
        foreach (int count in new[] { 100, 0 }) foreach (string language in new[] { "en", "pl" }) foreach (string format in new[] { "xlsx", "pdf" }) {
            await SelectHierarchyAsync(count, format, language);
            string contract = await ContractAsync(count, language);
            await GenerateAsync(page);
            Require(await page.Locator("#report-status").InnerTextAsync() == $"Ready. {count} services exported.", "Hierarchical completion count differs from the captured view.");
            files.Add(await DownloadAsync(page, output, $"hierarchical-{language}-{count}.{format}", count, contract));
        }
        await SelectHierarchyAsync(50000, "pdf", "pl");
        await page.Locator("#report-generate").ClickAsync();
        await page.WaitForFunctionAsync("document.querySelector('#report-progress').value > 0 && document.querySelector('#report-form').getAttribute('aria-busy') === 'true'", null, new() { Timeout = 120000 });
        double processed = await page.Locator("#report-progress").EvaluateAsync<double>("element => element.value");
        Require(await page.Locator("#report-language").IsDisabledAsync() && await page.Locator("#report-layout").IsDisabledAsync(), "Running export left mutable report controls enabled.");
        await page.Locator("#report-cancel").ClickAsync(); await ReadyAsync(page);
        Require(await page.Locator("#report-progress").EvaluateAsync<double>("element => element.value") == 0 && !await page.Locator("#report-download").IsVisibleAsync(), "Cancelled hierarchy exposed a partial download.");
        await SelectHierarchyAsync(100, "xlsx", "en"); await GenerateAsync(page);
        files.Add(await DownloadAsync(page, output, "hierarchical-recovery.xlsx", 100, await ContractAsync(100, "en")));
        await SelectAsync(page, 100, "csv");
        Require(await page.Locator("#report-layout").InputValueAsync() == "analysis" && await page.Locator("#report-layout").IsDisabledAsync() && await page.Locator("#report-language").IsDisabledAsync(), "CSV did not restore the flat export controls.");
        Require(await page.Locator("#report-preview tr").First.Locator("td").Nth(3).InnerTextAsync() == "4.8 ms", "CSV retained the hierarchical row source.");
        await GenerateAsync(page); files.Add(await DownloadAsync(page, output, "hierarchical-recovery-flat.csv", 100));
        states.Add(new { hierarchicalCancellationAfterRows = processed, partialDownloadCreated = false, xlsxRecovery = true, csvFlatRecovery = true });
        foreach (string theme in new[] { "light", "dark" }) {
            await page.EvaluateAsync("theme => { localStorage.setItem('imo-theme',theme); localStorage.setItem('imo-bg','blueprint'); }", theme);
            await page.ReloadAsync(); await ReadyAsync(page); await SelectHierarchyAsync(100, "pdf", "pl");
            foreach ((int width, int height, string label) in new[] { (1280, 900, "wide"), (390, 844, "phone") }) {
                await page.SetViewportSizeAsync(width, height); await page.EvaluateAsync("() => document.fonts.ready.then(() => true)");
                await page.Locator("#report-language").FocusAsync(); await page.Keyboard.PressAsync("Tab");
                if (await page.EvaluateAsync<string>("document.activeElement.id") != "report-generate" && engine == HtmlBrowserEngine.WebKit) {
                    await page.Locator("#report-language").FocusAsync(); await page.Keyboard.PressAsync("Alt+Tab");
                }
                Require(await page.EvaluateAsync<string>("document.activeElement.id") == "report-generate", "Hierarchical language control does not lead to Generate by keyboard.");
                Require(await page.EvaluateAsync<bool>("parseFloat(getComputedStyle(document.activeElement).outlineWidth) >= 2 && getComputedStyle(document.activeElement).outlineStyle !== 'none'"), "Hierarchical keyboard focus lacks a visible outline.");
                await ScrollTopAsync(page); JsonElement layout = await LayoutAsync(page);
                Require(layout.GetProperty("overflow").GetInt32() <= 1 && layout.GetProperty("controlsVisible").GetBoolean() && layout.GetProperty("theme").GetString() == theme, "Hierarchical layout overflow, controls or theme differ.");
                Require(layout.GetProperty("selectHeights").EnumerateArray().All(value => value.GetDouble() >= 44), "Hierarchical select controls are smaller than 44px.");
                await page.ScreenshotAsync(new() { Path = Path.Combine(output, $"hierarchical-{label}-{theme}.png"), FullPage = true });
                states.Add(new { theme, width, height, layout, hierarchicalLanguage = "pl", keyboardFocus = "language to generate" });
            }
        }
        await page.Locator("#report-rows").SelectOptionAsync("0"); await ScrollTopAsync(page);
        Require(await page.EvaluateAsync<bool>("document.querySelector('#report-preview').rows.length === 1 && document.querySelector('#report-preview td').colSpan === 4"), "Empty hierarchical preview lost its placeholder.");
        await page.ScreenshotAsync(new() { Path = Path.Combine(output, "hierarchical-phone-empty-dark.png"), FullPage = true });
    }
    private static Task<JsonElement> LayoutAsync(IPage page) => page.EvaluateAsync<JsonElement>("""
        () => ({ viewport: innerWidth, overflow: document.documentElement.scrollWidth - innerWidth,
          theme: document.documentElement.dataset.theme, background: document.documentElement.dataset.bg,
          controlsVisible: ['report-rows','report-format','report-layout','report-language','report-generate','report-cancel'].every(id => {
            const r = document.getElementById(id).getBoundingClientRect(); return r.width > 0 && r.left >= -1 && r.right <= innerWidth + 1;
          }), font: getComputedStyle(document.body).fontFamily,
          selectHeights: ['report-rows','report-format','report-layout','report-language'].map(id => document.getElementById(id).getBoundingClientRect().height),
          status: { live: document.querySelector('#report-status').getAttribute('aria-live'), busy: document.querySelector('#report-form').getAttribute('aria-busy') },
          previewScrollable: document.querySelector('.imo-report-lab__table-scroll').scrollWidth > document.querySelector('.imo-report-lab__table-scroll').clientWidth })
        """);
    private static async Task ScrollTopAsync(IPage page) {
        await page.EvaluateAsync("scrollTo({ top: 0, left: 0, behavior: 'instant' })");
        await page.WaitForFunctionAsync("scrollY === 0");
        await page.EvaluateAsync("() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))");
    }
    private static async Task<object> VerifyLifecycleAsync(IPage page, string output, bool expectDefect) {
        const string key = "officeimo-validation-lifecycle";
        await page.AddInitScriptAsync("""
            for (const type of ['pageshow', 'pagehide']) addEventListener(type, event => {
              const key = 'officeimo-validation-lifecycle';
              const events = JSON.parse(sessionStorage.getItem(key) || '[]');
              events.push({ type, persisted: event.persisted, trusted: event.isTrusted, path: location.pathname });
              sessionStorage.setItem(key, JSON.stringify(events));
            });
            """);
        await page.EvaluateAsync("key => sessionStorage.removeItem(key)", key);
        await page.ReloadAsync(); await ReadyAsync(page);
        await SelectAsync(page, 100, "csv"); await GenerateAsync(page);
        object before = await DownloadAsync(page, output, "lifecycle-before.csv", 100);
        await page.Locator(".imo-prose a[href='/convert/']").ClickAsync();
        await page.GoBackAsync(); await ReadyAsync(page);
        JsonElement actualEvents = await page.EvaluateAsync<JsonElement>("key => JSON.parse(sessionStorage.getItem(key) || '[]')", key);
        bool persisted = actualEvents.EnumerateArray().Any(e => e.GetProperty("type").GetString() == "pageshow" && e.GetProperty("persisted").GetBoolean() && e.GetProperty("trusted").GetBoolean());
        JsonElement actualDownload = await page.EvaluateAsync<JsonElement>("""
            async () => {
              const link = document.querySelector('#report-download');
              if (link.hidden) return { visible: false, blobFetchOk: false };
              try { const response = await fetch(link.href); return { visible: true, blobFetchOk: response.ok, bytes: (await response.arrayBuffer()).byteLength }; }
              catch (error) { return { visible: true, blobFetchOk: false, error: error.name }; }
            }
            """);
        if (persisted) Require(actualDownload.GetProperty("visible").GetBoolean() && actualDownload.GetProperty("blobFetchOk").GetBoolean() != expectDefect, "Actual BFCache restoration did not match the expected completed-download lifetime.");
        await SelectAsync(page, 100, "csv"); await GenerateAsync(page);
        JsonElement controlled = await page.EvaluateAsync<JsonElement>("""
            async () => {
              const link = document.querySelector('#report-download'), url = link.href;
              const before = await fetch(url); const bytesBefore = (await before.arrayBuffer()).byteLength;
              dispatchEvent(new PageTransitionEvent('pagehide', { persisted: true }));
              dispatchEvent(new PageTransitionEvent('pageshow', { persisted: true }));
              let fetchOk = false, bytesAfter = 0, error;
              try { const response = await fetch(url); fetchOk = response.ok; bytesAfter = (await response.arrayBuffer()).byteLength; }
              catch (caught) { error = caught.name; }
              return { beforeOk: before.ok, bytesBefore, afterOk: fetchOk, bytesAfter, error: error || null, visible: !link.hidden, sameHref: link.href === url };
            }
            """);
        Require(controlled.GetProperty("beforeOk").GetBoolean() && controlled.GetProperty("visible").GetBoolean() && controlled.GetProperty("sameHref").GetBoolean(), "Completed CSV was unavailable before persisted lifecycle transition.");
        Require(controlled.GetProperty("afterOk").GetBoolean() != expectDefect, "Controlled persisted lifecycle did not match the expected completed Blob lifetime.");
        object? preserved = expectDefect ? null : await DownloadAsync(page, output, "lifecycle-preserved.csv", 100);
        JsonElement finalDeparture = await page.EvaluateAsync<JsonElement>("""
            async () => { const url = document.querySelector('#report-download').href;
              dispatchEvent(new PageTransitionEvent('pagehide', { persisted: false }));
              try { await fetch(url); return { revoked: false }; } catch { return { revoked: true }; }
            }
            """);
        Require(finalDeparture.GetProperty("revoked").GetBoolean(), "Final nonpersisted departure retained the completed Blob URL.");
        return new { expectedDefect = expectDefect, actualNavigation = new { persistedRestorationObserved = persisted, events = actualEvents, download = actualDownload },
            controlledPersistedEvent = controlled, controlledFinalDeparture = finalDeparture, before, preserved,
            boundary = persisted ? "actual BFCache plus controlled PageTransitionEvent" : "actual navigate/back reloaded the document; persisted branch proven with controlled native PageTransitionEvent and real Blob fetch/download" };
    }
    private static void Require(bool condition, string message) => WorkbookVerifier.Require(condition, "Report workbench UI: " + message);
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
    private sealed record NetworkFailure(string Url, int Status);
}
