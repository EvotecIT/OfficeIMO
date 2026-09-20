using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class HtmlPublicBrowserEvidenceRunner {
    private const int MaximumActionManifestBytes = 256 * 1024;
    private const int MaximumActions = 64;
    private const int MaximumActionValues = 64;
    private const int MaximumInputCharacters = 64 * 1024;
    private const int MaximumReadyTimeoutMilliseconds = 60 * 1000;
    private static readonly TimeSpan MaximumBrowserRunTime = TimeSpan.FromMinutes(5);
    private static readonly TimeSpan MaximumBrowserCleanupTime = TimeSpan.FromSeconds(10);
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };
    private static readonly HashSet<string> KnownOptions = new(StringComparer.OrdinalIgnoreCase) {
        "--acquisition", "--officeimo-screen", "--output", "--document-url",
        "--viewport-width", "--viewport-height", "--actions", "--before-actions-expression",
        "--ready-expression", "--ready-timeout-ms"
    };

    internal static async Task<int> RunAsync(string[] args) {
        ValidateOptions(args);
        string acquisitionPath = RequiredPath(args, "--acquisition");
        string officeImoScreenPath = RequiredPath(args, "--officeimo-screen");
        string outputDirectory = Path.GetFullPath(Required(args, "--output"));
        Uri documentUrl = new(Required(args, "--document-url"), UriKind.Absolute);
        int viewportWidth = PositiveInt(args, "--viewport-width", 816);
        int viewportHeight = PositiveInt(args, "--viewport-height", 720);
        int readyTimeout = BoundedPositiveInt(args, "--ready-timeout-ms", 30000, MaximumReadyTimeoutMilliseconds);
        string? beforeActionsExpression = Optional(args, "--before-actions-expression");
        string? readyExpression = Optional(args, "--ready-expression");
        string? actionsPath = OptionalPath(args, "--actions");
        BrowserActionManifest? actionManifest = actionsPath == null ? null : ReadActions(actionsPath);
        IReadOnlyList<BrowserAction> actions = actionManifest?.Actions ?? Array.Empty<BrowserAction>();
        ValidateInputBudget(beforeActionsExpression, readyExpression, actions);
        if (Directory.Exists(outputDirectory) || File.Exists(outputDirectory))
            throw new IOException("The browser evidence output directory must not already exist.");

        Acquisition acquisition = ReadAcquisition(acquisitionPath);
        if (!acquisition.RetainedInputBytes)
            throw new InvalidDataException("Browser evidence requires acquisition with retained input bytes.");
        Dictionary<string, Resource> resources = LoadResources(acquisitionPath, acquisition);
        if (!resources.ContainsKey(documentUrl.AbsoluteUri))
            throw new InvalidDataException("The requested document URL is absent from the acquisition manifest.");

        byte[]? browserPng = null;
        string? browserVersion = null;
        DateTimeOffset deadline = DateTimeOffset.UtcNow + MaximumBrowserRunTime;
        using var runTimeout = new CancellationTokenSource(MaximumBrowserRunTime);
        var blockedUrls = new ConcurrentQueue<string>();
        HtmlBrowserSession? launcher = null;
        IBrowserContext? context = null;
        Exception? browserFailure = null;
        Exception? cleanupFailure = null;
        try {
            launcher = await HtmlPdfComparisonRenderers.OpenChromiumSessionAsync(runTimeout.Token)
                .WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
            IBrowser browserInstance = launcher.Browser
                ?? throw new InvalidOperationException("The browser-reference runner requires a non-persistent browser instance.");
            browserVersion = browserInstance.Version;
            context = await browserInstance.NewContextAsync(new BrowserNewContextOptions {
                ServiceWorkers = ServiceWorkerPolicy.Block,
                ViewportSize = new ViewportSize { Width = viewportWidth, Height = viewportHeight }
            }).WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
            await context.RouteAsync("**/*", async route => {
                IRequest request = route.Request;
                if (!string.Equals(request.Method, "GET", StringComparison.OrdinalIgnoreCase)
                    || request.PostDataBuffer is { Length: > 0 }) {
                    blockedUrls.Enqueue(request.Method + " " + request.Url);
                    await route.AbortAsync("blockedbyclient").ConfigureAwait(false);
                    return;
                }
                if (resources.TryGetValue(request.Url, out Resource? resource)) {
                    await route.FulfillAsync(new RouteFulfillOptions {
                        BodyBytes = resource.Bytes,
                        ContentType = resource.ContentType,
                        Status = resource.StatusCode
                    }).ConfigureAwait(false);
                    return;
                }
                blockedUrls.Enqueue(request.Method + " " + request.Url);
                await route.AbortAsync("blockedbyclient").ConfigureAwait(false);
            }).WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
            IPage page = await context.NewPageAsync()
                .WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
            await page.GotoAsync(documentUrl.AbsoluteUri, new PageGotoOptions { WaitUntil = WaitUntilState.Load })
                .WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
            await WaitForExpressionAsync(page, beforeActionsExpression, OperationTimeout(readyTimeout, deadline))
                .WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
            foreach (BrowserAction action in actions) {
                await RunActionAsync(page, action, OperationTimeout(readyTimeout, deadline))
                    .WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
            }
            await WaitForExpressionAsync(page, readyExpression, OperationTimeout(readyTimeout, deadline))
                .WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
            await page.EvaluateAsync("document.fonts ? document.fonts.ready : Promise.resolve()")
                .WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
            browserPng = await page.ScreenshotAsync(new PageScreenshotOptions {
                FullPage = true,
                Type = ScreenshotType.Png,
                Animations = ScreenshotAnimations.Disabled,
                Caret = ScreenshotCaret.Hide
            }).WaitAsync(RemainingRunTime(deadline), runTimeout.Token).ConfigureAwait(false);
        } catch (Exception exception) {
            browserFailure = exception;
        } finally {
            cleanupFailure = await DisposeBrowserResourcesAsync(context, launcher).ConfigureAwait(false);
        }
        if (browserFailure != null && cleanupFailure != null) throw new AggregateException(browserFailure, cleanupFailure);
        if (browserFailure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(browserFailure).Throw();
        if (cleanupFailure != null) throw cleanupFailure;
        if (browserPng == null) throw new InvalidOperationException("The browser-reference runner completed without a screenshot.");
        if (browserVersion == null) throw new InvalidOperationException("The browser-reference runner completed without a browser version.");
        if (!blockedUrls.IsEmpty)
            throw new InvalidDataException("The offline browser reference requested resources absent from the acquisition: " +
                string.Join(", ", blockedUrls.Distinct(StringComparer.Ordinal)));

        byte[] officeImoPng = File.ReadAllBytes(officeImoScreenPath);
        OfficeRasterImage browser = Decode(browserPng, "Chromium screenshot");
        OfficeRasterImage officeImo = Decode(officeImoPng, "OfficeIMO screenshot");
        PixelComparison comparison = Compare(browser, officeImo);
        byte[] differencePng = OfficePngWriter.Encode(comparison.Difference);
        Directory.CreateDirectory(outputDirectory);
        File.WriteAllBytes(Path.Combine(outputDirectory, "chromium-screen.png"), browserPng);
        File.WriteAllBytes(Path.Combine(outputDirectory, "officeimo-screen.png"), officeImoPng);
        File.WriteAllBytes(Path.Combine(outputDirectory, "difference.png"), differencePng);
        File.WriteAllText(Path.Combine(outputDirectory, "browser-evidence.json"), JsonSerializer.Serialize(new {
            schemaVersion = 1,
            scenarioId = acquisition.ScenarioId,
            sourceLicense = acquisition.SourceLicense,
            documentUrl = documentUrl.AbsoluteUri,
            capturedAtUtc = DateTimeOffset.UtcNow,
            browser = new {
                product = "Chromium",
                version = browserVersion,
                playwrightAssemblyVersion = typeof(IPlaywright).Assembly.GetName().Version?.ToString() ?? "unknown",
                htmlTinkerXAssemblyVersion = typeof(HtmlBrowser).Assembly.GetName().Version?.ToString() ?? "unknown",
                hostOperatingSystem = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
                hostArchitecture = System.Runtime.InteropServices.RuntimeInformation.OSArchitecture.ToString(),
                framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription
            },
            viewport = new { width = viewportWidth, height = viewportHeight },
            acquisition = new {
                file = Path.GetFileName(acquisitionPath),
                sha256 = Sha256(File.ReadAllBytes(acquisitionPath)),
                retainedInputBytes = true,
                networkPolicy = "context-route-all;service-workers-blocked;fulfill-static-get-acquired-bytes;abort-unknown"
            },
            finalState = new {
                beforeActionsExpression,
                readyExpression,
                readyTimeoutMs = readyTimeout,
                actionManifest = actionsPath == null ? null : new {
                    file = Path.GetFileName(actionsPath),
                    sha256 = actionManifest!.Sha256,
                    count = actions.Count
                }
            },
            source = resources.Values.Distinct().Select(resource => new {
                url = resource.Url,
                resource.ContentType,
                resource.StatusCode,
                bytes = resource.Bytes.LongLength,
                sha256 = Sha256(resource.Bytes)
            }).OrderBy(resource => resource.url, StringComparer.Ordinal).ToArray(),
            officeImo = new {
                file = "officeimo-screen.png",
                width = officeImo.Width,
                height = officeImo.Height,
                bytes = officeImoPng.LongLength,
                sha256 = Sha256(officeImoPng)
            },
            chromium = new {
                file = "chromium-screen.png",
                width = browser.Width,
                height = browser.Height,
                bytes = browserPng.LongLength,
                sha256 = Sha256(browserPng)
            },
            comparison = new {
                dimensionsMatch = browser.Width == officeImo.Width && browser.Height == officeImo.Height,
                comparison.Width,
                comparison.Height,
                comparison.MeanAbsoluteError,
                comparison.RootMeanSquareError,
                comparison.MeanLuminanceError,
                differenceFile = "difference.png",
                differenceSha256 = Sha256(differencePng)
            },
            blockedUrls = Array.Empty<string>()
        }, JsonOptions));
        Console.WriteLine("HTML_PUBLIC_BROWSER_EVIDENCE=" + Path.Combine(outputDirectory, "browser-evidence.json"));
        return 0;
    }

    private static BrowserActionManifest ReadActions(string path) {
        byte[] bytes = ReadBoundedFile(path, MaximumActionManifestBytes);
        BrowserAction[] actions = JsonSerializer.Deserialize<BrowserAction[]>(bytes, new JsonSerializerOptions {
            PropertyNameCaseInsensitive = true
        }) ?? throw new InvalidDataException("The browser action manifest is empty.");
        if (actions.Length > MaximumActions)
            throw new InvalidDataException($"The browser action manifest exceeds the {MaximumActions}-action limit.");
        foreach (BrowserAction action in actions) {
            if (string.IsNullOrWhiteSpace(action.Action) || string.IsNullOrWhiteSpace(action.Selector))
                throw new InvalidDataException("Every browser action requires action and selector.");
            string name = action.Action.Trim().ToLowerInvariant();
            if (name is not ("fill" or "select" or "set-checked" or "click" or "wait-for-text"))
                throw new InvalidDataException("Unsupported browser evidence action: " + action.Action);
            if (name is "fill" or "wait-for-text" && action.Value == null)
                throw new InvalidDataException(action.Action + " requires value.");
            if (name == "set-checked" && action.Checked == null)
                throw new InvalidDataException("set-checked requires checked.");
            if (name == "select" && (action.Values == null || action.Values.Count == 0))
                throw new InvalidDataException("select requires at least one value.");
            if (action.Values is { Count: > MaximumActionValues })
                throw new InvalidDataException($"A browser select action exceeds the {MaximumActionValues}-value limit.");
        }
        return new BrowserActionManifest(Array.AsReadOnly(actions), Sha256(bytes));
    }

    private static async Task<Exception?> DisposeBrowserResourcesAsync(IBrowserContext? context, HtmlBrowserSession? launcher) {
        var failures = new List<Exception>(2);
        if (context != null) {
            try {
                await context.DisposeAsync().AsTask().WaitAsync(MaximumBrowserCleanupTime).ConfigureAwait(false);
            } catch (Exception exception) {
                failures.Add(exception);
            }
        }
        if (launcher != null) {
            try {
                await launcher.DisposeAsync().AsTask().WaitAsync(MaximumBrowserCleanupTime).ConfigureAwait(false);
            } catch (Exception exception) {
                failures.Add(exception);
            }
        }
        return failures.Count switch {
            0 => null,
            1 => failures[0],
            _ => new AggregateException("Browser evidence cleanup failed.", failures)
        };
    }

    private static byte[] ReadBoundedFile(string path, int maximumBytes) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > maximumBytes)
            throw new InvalidDataException($"The browser action manifest exceeds the {maximumBytes}-byte limit.");
        var bytes = new byte[checked((int)stream.Length)];
        stream.ReadExactly(bytes);
        if (stream.ReadByte() >= 0)
            throw new InvalidDataException($"The browser action manifest exceeds the {maximumBytes}-byte limit.");
        return bytes;
    }

    private static void ValidateInputBudget(string? beforeActionsExpression, string? readyExpression, IReadOnlyList<BrowserAction> actions) {
        long characters = (beforeActionsExpression?.Length ?? 0) + (readyExpression?.Length ?? 0);
        foreach (BrowserAction action in actions) {
            characters += action.Action?.Length ?? 0;
            characters += action.Selector?.Length ?? 0;
            characters += action.Value?.Length ?? 0;
            if (action.Values != null) characters += action.Values.Sum(value => (long)(value?.Length ?? 0));
        }
        if (characters > MaximumInputCharacters)
            throw new InvalidDataException($"Browser readiness and action inputs exceed the {MaximumInputCharacters}-character limit.");
    }

    private static int OperationTimeout(int requestedMilliseconds, DateTimeOffset deadline) {
        double remaining = (deadline - DateTimeOffset.UtcNow).TotalMilliseconds;
        if (remaining <= 0D) throw new TimeoutException("The browser evidence run exceeded its total time limit.");
        return Math.Max(1, Math.Min(requestedMilliseconds, (int)Math.Min(int.MaxValue, remaining)));
    }

    private static TimeSpan RemainingRunTime(DateTimeOffset deadline) {
        TimeSpan remaining = deadline - DateTimeOffset.UtcNow;
        if (remaining <= TimeSpan.Zero) throw new TimeoutException("The browser evidence run exceeded its total time limit.");
        return remaining;
    }

    private static async Task RunActionAsync(IPage page, BrowserAction action, int timeout) {
        ILocator locator = page.Locator(action.Selector);
        switch (action.Action.Trim().ToLowerInvariant()) {
            case "fill":
                await locator.FillAsync(action.Value!, new LocatorFillOptions { Timeout = timeout }).ConfigureAwait(false);
                break;
            case "select":
                await locator.SelectOptionAsync(action.Values!.Select(value => new SelectOptionValue { Value = value }).ToArray(),
                    new LocatorSelectOptionOptions { Timeout = timeout }).ConfigureAwait(false);
                break;
            case "set-checked":
                await locator.SetCheckedAsync(action.Checked!.Value, new LocatorSetCheckedOptions { Timeout = timeout }).ConfigureAwait(false);
                break;
            case "click":
                await locator.ClickAsync(new LocatorClickOptions { Timeout = timeout }).ConfigureAwait(false);
                break;
            case "wait-for-text":
                await page.WaitForFunctionAsync(
                    "([selector, value]) => document.querySelector(selector)?.textContent === value",
                    new[] { action.Selector, action.Value! },
                    new PageWaitForFunctionOptions { Timeout = timeout }).ConfigureAwait(false);
                break;
        }
    }

    private static Task WaitForExpressionAsync(IPage page, string? expression, int timeout) =>
        string.IsNullOrWhiteSpace(expression)
            ? Task.CompletedTask
            : page.WaitForFunctionAsync(expression, null, new PageWaitForFunctionOptions { Timeout = timeout });

    private static Acquisition ReadAcquisition(string path) {
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(path));
        JsonElement root = json.RootElement;
        return new Acquisition(
            root.GetProperty("ScenarioId").GetString() ?? throw new InvalidDataException("ScenarioId is empty."),
            root.GetProperty("SourceLicense").GetString() ?? throw new InvalidDataException("SourceLicense is empty."),
            root.GetProperty("retainedInputBytes").GetBoolean(),
            root.GetProperty("resources").EnumerateArray().Select(resource => new ResourceManifest(
                resource.GetProperty("url").GetString() ?? string.Empty,
                resource.GetProperty("finalUrl").GetString() ?? string.Empty,
                resource.GetProperty("ContentType").GetString() ?? "application/octet-stream",
                resource.GetProperty("StatusCode").GetInt32(),
                resource.GetProperty("Sha256").GetString() ?? string.Empty,
                resource.GetProperty("RequestMethod").GetString() ?? string.Empty,
                resource.GetProperty("RequestOccurrence").ValueKind == JsonValueKind.Null
                    ? null
                    : resource.GetProperty("RequestOccurrence").GetInt32(),
                resource.GetProperty("dynamicExchanges").GetArrayLength(),
                resource.GetProperty("redirects").GetArrayLength())).ToArray());
    }

    private static Dictionary<string, Resource> LoadResources(string acquisitionPath, Acquisition acquisition) {
        string inputs = Path.Combine(Path.GetDirectoryName(acquisitionPath)!, "inputs");
        var result = new Dictionary<string, Resource>(StringComparer.Ordinal);
        foreach (ResourceManifest item in acquisition.Resources) {
            if (!string.Equals(item.RequestMethod, "GET", StringComparison.OrdinalIgnoreCase)
                || item.RequestOccurrence.HasValue
                || item.DynamicExchangeCount != 0
                || item.RedirectCount != 0
                || !string.Equals(item.Url, item.FinalUrl, StringComparison.Ordinal)) {
                throw new InvalidDataException(
                    "The browser-reference runner accepts only direct static GET acquisitions. " +
                    "Dynamic requests and redirect chains require exact request/response exchange bytes.");
            }
            string path = Path.Combine(inputs, item.Sha256);
            byte[] bytes = File.ReadAllBytes(path);
            if (!string.Equals(Sha256(bytes), item.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Retained input does not match its acquisition digest: " + item.Url);
            var resource = new Resource(item.Url, item.ContentType, item.StatusCode, bytes);
            if (!result.TryAdd(item.Url, resource))
                throw new InvalidDataException("The static acquisition contains a duplicate URL identity: " + item.Url);
        }
        return result;
    }

    private static PixelComparison Compare(OfficeRasterImage expected, OfficeRasterImage actual) {
        int width = Math.Min(expected.Width, actual.Width);
        int height = Math.Min(expected.Height, actual.Height);
        if (width < 1 || height < 1) throw new InvalidDataException("Screenshots have no comparable area.");
        long absolute = 0;
        long squared = 0;
        double luminance = 0D;
        var difference = new OfficeRasterImage(width, height, OfficeColor.White);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                OfficeColor left = expected.GetPixel(x, y);
                OfficeColor right = actual.GetPixel(x, y);
                int red = Math.Abs(left.R - right.R);
                int green = Math.Abs(left.G - right.G);
                int blue = Math.Abs(left.B - right.B);
                int alpha = Math.Abs(left.A - right.A);
                absolute += red + green + blue + alpha;
                squared += red * red + green * green + blue * blue + alpha * alpha;
                luminance += Math.Abs(left.R * 0.2126D + left.G * 0.7152D + left.B * 0.0722D -
                    (right.R * 0.2126D + right.G * 0.7152D + right.B * 0.0722D));
                int maximum = Math.Max(red, Math.Max(green, Math.Max(blue, alpha)));
                difference.SetPixel(x, y, OfficeColor.FromRgb((byte)Math.Min(255, maximum * 5), 0, 0));
            }
        }
        double channels = width * height * 4D;
        double pixels = width * height;
        return new PixelComparison(width, height, absolute / channels,
            Math.Sqrt(squared / channels), luminance / pixels, difference);
    }

    private static OfficeRasterImage Decode(byte[] png, string name) =>
        OfficePngReader.TryDecode(png, out OfficeRasterImage? image) && image != null
            ? image
            : throw new InvalidDataException(name + " is not a supported PNG image.");

    private static string RequiredPath(string[] args, string name) {
        string path = Path.GetFullPath(Required(args, name));
        return File.Exists(path) ? path : throw new FileNotFoundException(name + " does not exist.", path);
    }

    private static string? OptionalPath(string[] args, string name) {
        string? value = Optional(args, name);
        if (value == null) return null;
        string path = Path.GetFullPath(value);
        return File.Exists(path) ? path : throw new FileNotFoundException(name + " does not exist.", path);
    }

    private static void ValidateOptions(string[] args) {
        if (args.Length < 2 || !string.Equals(args[0], "html-public-browser-evidence", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("The browser evidence command is missing.");
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int index = 1; index < args.Length; index += 2) {
            string option = args[index];
            if (!KnownOptions.Contains(option))
                throw new ArgumentException("Unknown browser evidence option " + option + ".");
            if (!seen.Add(option))
                throw new ArgumentException("Duplicate browser evidence option " + option + ".");
            if (index + 1 >= args.Length || string.IsNullOrWhiteSpace(args[index + 1]) ||
                args[index + 1].StartsWith("--", StringComparison.Ordinal))
                throw new ArgumentException("Missing value for browser evidence option " + option + ".");
        }
    }

    private static string Required(string[] args, string name) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length && !string.IsNullOrWhiteSpace(args[index + 1])
            ? args[index + 1]
            : throw new ArgumentException("Missing required option " + name + ".");
    }

    private static string? Optional(string[] args, string name) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static int PositiveInt(string[] args, string name, int fallback) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return fallback;
        return index + 1 < args.Length && int.TryParse(args[index + 1], out int value) && value > 0
            ? value
            : throw new ArgumentException(name + " must be a positive integer.");
    }

    private static int BoundedPositiveInt(string[] args, string name, int fallback, int maximum) {
        int value = PositiveInt(args, name, fallback);
        return value <= maximum
            ? value
            : throw new ArgumentOutOfRangeException(name, $"{name} must not exceed {maximum}.");
    }

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private sealed record Acquisition(string ScenarioId, string SourceLicense, bool RetainedInputBytes,
        IReadOnlyList<ResourceManifest> Resources);
    private sealed record ResourceManifest(string Url, string FinalUrl, string ContentType, int StatusCode, string Sha256,
        string RequestMethod, int? RequestOccurrence, int DynamicExchangeCount, int RedirectCount);
    private sealed record Resource(string Url, string ContentType, int StatusCode, byte[] Bytes);
    private sealed record BrowserActionManifest(IReadOnlyList<BrowserAction> Actions, string Sha256);
    private sealed record BrowserAction(string Action, string Selector, string? Value, IReadOnlyList<string>? Values, bool? Checked);
    private sealed record PixelComparison(int Width, int Height, double MeanAbsoluteError,
        double RootMeanSquareError, double MeanLuminanceError, OfficeRasterImage Difference);
}
