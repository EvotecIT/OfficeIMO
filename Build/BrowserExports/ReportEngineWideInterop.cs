using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;

/// <summary>Exercises the real wide-report worker and the completed actions at the affected wrapping boundaries.</summary>
internal static partial class ReportEngineInterop {
    private static async Task VerifyWideAsync(IPage page, HtmlBrowserEngine engine, string output, List<object> files, List<object> states) {
        async Task SelectWideAsync(int count, string format, string language) {
            await SelectAsync(page, count, format);
            await page.Locator("#report-layout").SelectOptionAsync("wide");
            await page.Locator("#report-language").SelectOptionAsync(language);
            Require(await page.Locator("#report-language").IsEnabledAsync() && await page.Locator("#report-preview-scope").IsVisibleAsync(), "Wide preview/language context is unavailable.");
            Require(!await page.Locator("#report-download").IsVisibleAsync() && !await page.Locator("#report-cancel").IsVisibleAsync(), "Changed report settings retained a stale or inactive action.");
        }
        async Task<string> ContractAsync(int count, string language) => await page.EvaluateAsync<string>("""
            async ({count,language}) => {
              const { wideAnalyticalRows, wideAnalyticalColumns, wideAnalyticalText } = await import('/js/wide-analytical-report-job.mjs');
              const { analyticalReportText } = await import('/js/analytical-report-job.mjs');
              const text = analyticalReportText[language], extra = wideAnalyticalText[language];
              const columns = wideAnalyticalColumns.map(column => ({...column,header: text.headers[column.key] ?? extra.headers[column.key] ?? column.header}));
              const decimal = new Intl.NumberFormat(language,{maximumFractionDigits:3}), percent = new Intl.NumberFormat(language,{style:'percent',maximumFractionDigits:2});
              const percentages = new Set(['availability','errorRate','cpu']);
              const records = [...wideAnalyticalRows(count)].map(row => row.slice(0,columns.length));
              return JSON.stringify({language,columns,rows:records,scope: language === 'pl'
                ? `Wybrany widok usług · ${count.toLocaleString('pl')} rekordów.` : `Selected service view · ${count.toLocaleString('en')} services.`,
                title:text.title,ending:text.next,advice:text.advice,page:text.page,
                display:records.map(row=>row.map((value,index)=>value instanceof Date ? value.toISOString().slice(0,16).replace('T',' ')
                  : typeof value === 'number' ? (percentages.has(columns[index].key) ? percent : decimal).format(value) : String(value ?? ''))) });
            }
            """, new { count, language });
        async Task SaveAsync(int count, string format, string language) {
            string name = $"wide-{language}-{count}.{format}", path = Path.Combine(output, name), contract = await ContractAsync(count, language);
            var download = await page.RunAndWaitForDownloadAsync(() => page.Locator("#report-download").ClickAsync(), new() { Timeout = 120000 });
            await download.SaveAsAsync(path);
            Require(await download.FailureAsync() is null && download.SuggestedFilename == "service-health." + format, "Wide report download failed or used the wrong filename.");
            File.WriteAllText(path + ".json", contract); files.Add(ReportEngineWideVerifier.Verify(path));
        }
        foreach (int count in new[] { 100, 0 }) foreach (string language in new[] { "en", "pl" }) foreach (string format in new[] { "xlsx", "pdf" }) {
            await SelectWideAsync(count, format, language); await GenerateAsync(page);
            Require(await page.Locator("#report-status").InnerTextAsync() == $"Ready. {count} services exported.", "Wide completion count differs from its selected source.");
            await VerifyCompletedActionsAsync(page); await SaveAsync(count, format, language);
        }
        await page.Locator("#report-generate").FocusAsync(); await page.Keyboard.PressAsync("Shift+Tab");
        if (await page.EvaluateAsync<string>("document.activeElement.id") != "report-download" && engine == HtmlBrowserEngine.WebKit) {
            await page.Locator("#report-generate").FocusAsync(); await page.Keyboard.PressAsync("Alt+Shift+Tab");
        }
        Require(await page.EvaluateAsync<string>("document.activeElement.id") == "report-download", "The completed primary download is not reachable immediately before Generate again by keyboard.");
        await SelectWideAsync(50000, "pdf", "pl"); await page.Locator("#report-generate").ClickAsync();
        await page.WaitForFunctionAsync("document.querySelector('#report-progress').value > 0 && document.querySelector('#report-form').getAttribute('aria-busy') === 'true'", null, new() { Timeout = 120000 });
        double processed = await page.Locator("#report-progress").EvaluateAsync<double>("element=>element.value");
        Require(await page.Locator("#report-cancel").IsVisibleAsync() && await page.Locator("#report-language").IsDisabledAsync(), "Running wide report lost its cancellation or froze incomplete settings.");
        await page.BringToFrontAsync(); await page.Locator("#report-cancel").FocusAsync();
        Require(await page.EvaluateAsync<bool>("document.hasFocus() && document.activeElement.id === 'report-cancel'"), "Cancellation focus precondition is not observable.");
        await page.Keyboard.PressAsync("Enter"); await ReadyAsync(page);
        Require(!await page.Locator("#report-download").IsVisibleAsync() && !await page.Locator("#report-cancel").IsVisibleAsync() && await page.Locator("#report-progress").EvaluateAsync<double>("element=>element.value") == 0, "Cancelled wide report exposed partial output or inactive cancellation.");
        Require(await page.EvaluateAsync<string>("document.activeElement.id") == "report-generate", "Hiding the focused Cancel lost keyboard focus.");
        await page.Locator("#report-generate").ClickAsync();
        await page.WaitForFunctionAsync("document.querySelector('#report-progress').value > 0 && document.querySelector('#report-form').getAttribute('aria-busy') === 'true'", null, new() { Timeout = 120000 });
        await page.EvaluateAsync("""
            () => { const cancel = document.querySelector('#report-cancel'); cancel.focus(); cancel.click(); document.querySelector('.imo-report-lab__table-scroll').focus(); }
            """);
        await ReadyAsync(page);
        Require(await page.EvaluateAsync<bool>("document.activeElement.classList.contains('imo-report-lab__table-scroll')"), "Cancelled work stole focus from a user-selected reachable control.");
        await SelectAsync(page, 100, "csv");
        Require(await page.Locator("#report-layout").InputValueAsync() == "analysis" && await page.Locator("#report-language").IsDisabledAsync() && !await page.Locator("#report-preview-scope").IsVisibleAsync(), "CSV did not restore the canonical flat report.");
        await GenerateAsync(page); files.Add(await DownloadAsync(page, output, "wide-recovery-flat.csv", 100));
        states.Add(new { cancellationAfterRows = processed, partialDownloadCreated = false, flatCsvRecovery = true,
            cancellationKeyboardFocusRestored = true, completedDownloadKeyboardReachable = true, controlledSameTaskFocusMovePreserved = true });
        foreach (string theme in new[] { "light", "dark" }) {
            await page.EvaluateAsync("theme=>{localStorage.setItem('imo-theme',theme);localStorage.setItem('imo-bg','blueprint');}", theme);
            await page.ReloadAsync(); await ReadyAsync(page); await SelectWideAsync(100, "pdf", "pl"); await GenerateAsync(page);
            foreach ((int width, int height, string label) in new[] { (1440, 1000, "wide"), (1024, 768, "sidebar"), (862, 650, "intermediate-short"), (390, 844, "phone") }) {
                await page.SetViewportSizeAsync(width, height); await page.EvaluateAsync("()=>document.fonts.ready.then(()=>true)");
                await SettleViewportAsync(page);
                JsonElement measurement = await VerifyCompletedActionsAsync(page), layout = await LayoutAsync(page);
                Require(layout.GetProperty("overflow").GetInt32() <= 1 && layout.GetProperty("controlsVisible").GetBoolean(), "Completed wide report overflowed or lost controls.");
                await page.Locator("#report-download").FocusAsync();
                Require(await page.EvaluateAsync<bool>("parseFloat(getComputedStyle(document.activeElement).outlineWidth)>=2"), "Completed download lacks visible keyboard focus.");
                await page.Locator(".imo-report-lab__actions").ScrollIntoViewIfNeededAsync();
                await SettleViewportAsync(page);
                await page.ScreenshotAsync(new() { Path = Path.Combine(output, $"wide-completed-{label}-{theme}.png"), FullPage = true });
                await page.Locator(".imo-report-lab__actions").ScreenshotAsync(new() { Path = Path.Combine(output, $"wide-actions-{label}-{theme}.png") });
                await page.Locator(".imo-report-lab__result").ScreenshotAsync(new() { Path = Path.Combine(output, $"wide-result-{label}-{theme}.png") });
                states.Add(new { theme, width, height, layout, measurement, completedActions = true, primaryDownload = true, finiteTransitionsSettled = true, reportLanguage = "pl" });
            }
        }
    }
}
