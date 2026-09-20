using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class HtmlPublicBrowserEvidenceRunner {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };
    private static readonly HashSet<string> KnownOptions = new(StringComparer.OrdinalIgnoreCase) {
        "--acquisition", "--officeimo-screen", "--output", "--document-url",
        "--viewport-width", "--viewport-height"
    };

    internal static async Task<int> RunAsync(string[] args) {
        ValidateOptions(args);
        string acquisitionPath = RequiredPath(args, "--acquisition");
        string officeImoScreenPath = RequiredPath(args, "--officeimo-screen");
        string outputDirectory = Path.GetFullPath(Required(args, "--output"));
        Uri documentUrl = new(Required(args, "--document-url"), UriKind.Absolute);
        int viewportWidth = PositiveInt(args, "--viewport-width", 816);
        int viewportHeight = PositiveInt(args, "--viewport-height", 720);
        if (Directory.Exists(outputDirectory) || File.Exists(outputDirectory))
            throw new IOException("The browser evidence output directory must not already exist.");

        Acquisition acquisition = ReadAcquisition(acquisitionPath);
        if (!acquisition.RetainedInputBytes)
            throw new InvalidDataException("Browser evidence requires acquisition with retained input bytes.");
        Dictionary<string, Resource> resources = LoadResources(acquisitionPath, acquisition);
        if (!resources.ContainsKey(documentUrl.AbsoluteUri))
            throw new InvalidDataException("The requested document URL is absent from the acquisition manifest.");

        byte[] browserPng;
        string browserVersion;
        var blockedUrls = new ConcurrentQueue<string>();
        await using (HtmlBrowserSession launcher = await HtmlPdfComparisonRenderers.OpenChromiumSessionAsync().ConfigureAwait(false)) {
            IBrowser browserInstance = launcher.Browser
                ?? throw new InvalidOperationException("The browser-reference runner requires a non-persistent browser instance.");
            browserVersion = browserInstance.Version;
            await using IBrowserContext context = await browserInstance.NewContextAsync(new BrowserNewContextOptions {
                ServiceWorkers = ServiceWorkerPolicy.Block,
                ViewportSize = new ViewportSize { Width = viewportWidth, Height = viewportHeight }
            }).ConfigureAwait(false);
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
            }).ConfigureAwait(false);
            IPage page = await context.NewPageAsync().ConfigureAwait(false);
            await page.GotoAsync(documentUrl.AbsoluteUri, new PageGotoOptions { WaitUntil = WaitUntilState.Load }).ConfigureAwait(false);
            await page.EvaluateAsync("document.fonts ? document.fonts.ready : Promise.resolve()").ConfigureAwait(false);
            browserPng = await page.ScreenshotAsync(new PageScreenshotOptions {
                FullPage = true,
                Type = ScreenshotType.Png,
                Animations = ScreenshotAnimations.Disabled,
                Caret = ScreenshotCaret.Hide
            }).ConfigureAwait(false);
        }
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

    private static int PositiveInt(string[] args, string name, int fallback) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return fallback;
        return index + 1 < args.Length && int.TryParse(args[index + 1], out int value) && value > 0
            ? value
            : throw new ArgumentException(name + " must be a positive integer.");
    }

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private sealed record Acquisition(string ScenarioId, string SourceLicense, bool RetainedInputBytes,
        IReadOnlyList<ResourceManifest> Resources);
    private sealed record ResourceManifest(string Url, string FinalUrl, string ContentType, int StatusCode, string Sha256,
        string RequestMethod, int? RequestOccurrence, int DynamicExchangeCount, int RedirectCount);
    private sealed record Resource(string Url, string ContentType, int StatusCode, byte[] Bytes);
    private sealed record PixelComparison(int Width, int Height, double MeanAbsoluteError,
        double RootMeanSquareError, double MeanLuminanceError, OfficeRasterImage Difference);
}
