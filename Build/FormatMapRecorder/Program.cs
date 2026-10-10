using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Sockets;
using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;

/// <summary>
/// Renders the website's format-map stage (/format-map/) into marketing cuts: one clip per entry of cuts.json, framed for its
/// platform (16:9, 1:1, 4:5, 9:16), plus a poster frame.
///
/// The render is frame by frame, not a screen recording. The built site is served from a private loopback server, the page clock
/// is frozen with Playwright's fake clock, and each frame advances timers and requestAnimationFrame by exactly 1/fps seconds, steps
/// the CSS animations and transitions by the same amount, then takes a screenshot at the cut's device scale. The result is the same
/// at any speed of machine, any frame rate and any resolution (--scale=2 renders 1920x1080 layouts as true 3840x2160), and nothing
/// has to be trimmed. ffmpeg then encodes the frames.
///
///   dotnet run --project Build/FormatMapRecorder -- [--site=Website/_site] [--output=Artefacts/FormatMapVideos]
///                                                    [--cut=name,name] [--scale=2] [--fps=60] [--formats=mp4,mov,...]
///                                                    [--cuts=file] [--ffmpeg=path] [--list]
/// Output formats: mp4 (H.264), h265, av1, webm (VP9), mov (ProRes 422 HQ), gif, webp, apng.
/// Frames are written under EVOTEC_SCRATCH_ROOT (or the output folder) and removed after a successful encode.
/// </summary>
internal static class Program {
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true, ReadCommentHandling = JsonCommentHandling.Skip };

    // Advances every CSS animation and transition by dt. A new one is paused the moment it is first seen, so no real time passes in it
    // while a screenshot is taken.
    private const string StepScript = """
        (dt) => {
          document.body.getBoundingClientRect();
          for (const animation of document.getAnimations()) {
            if (!animation.__imoSeen) { animation.__imoSeen = true; animation.pause(); animation.currentTime = 0; }
            else if (dt) { animation.currentTime = animation.currentTime + dt; }
          }
        }
        """;

    private static async Task<int> Main(string[] args) {
        string root = Directory.GetCurrentDirectory();
        string site = Path.GetFullPath(Option(args, "--site") ?? Path.Combine(root, "Website", "_site"));
        string output = Path.GetFullPath(Option(args, "--output") ?? Path.Combine(root, "Artefacts", "FormatMapVideos"));
        string cutsPath = Option(args, "--cuts") ?? Path.Combine(AppContext.BaseDirectory, "cuts.json");
        Cut[] all = JsonSerializer.Deserialize<CutFile>(File.ReadAllText(cutsPath), JsonOptions)?.Cuts ?? [];

        if (args.Contains("--list", StringComparer.OrdinalIgnoreCase)) {
            foreach (Cut cut in all) Console.WriteLine($"{cut.Name,-24} {cut.Width}x{cut.Height}  {string.Join('+', cut.Formats)}  {cut.Scenes ?? "(all scenes)"}");
            return 0;
        }

        string[] wanted = (Option(args, "--cut") ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        string[] unknown = wanted.Where(name => !all.Any(cut => string.Equals(cut.Name, name, StringComparison.OrdinalIgnoreCase))).ToArray();
        if (unknown.Length > 0) {
            Console.Error.WriteLine("Unknown cut: " + string.Join(", ", unknown));
            return 2;
        }
        Cut[] cuts = wanted.Length == 0 ? all : all.Where(cut => wanted.Contains(cut.Name, StringComparer.OrdinalIgnoreCase)).ToArray();
        // --scale=2 renders the same layout at twice the pixels (1920x1080 becomes 3840x2160); --formats replaces each cut's formats.
        int scale = int.TryParse(Option(args, "--scale"), out int parsedScale) ? Math.Clamp(parsedScale, 1, 4) : 1;
        int? fps = int.TryParse(Option(args, "--fps"), out int parsedFps) ? Math.Clamp(parsedFps, 10, 120) : null;
        string[]? formats = Option(args, "--formats")?.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        cuts = cuts.Select(cut => cut with {
            Name = scale > 1 ? $"{cut.Name}-{scale}x" : cut.Name,
            Width = cut.Width * scale, Height = cut.Height * scale, Scale = cut.Scale * scale,
            Fps = fps ?? cut.Fps,
            Formats = formats is { Length: > 0 } ? formats : cut.Formats
        }).ToArray();

        if (!File.Exists(Path.Combine(site, "format-map", "index.html"))) {
            Console.Error.WriteLine($"No built format-map page under {site}. Build the website first (Website/build.ps1).");
            return 2;
        }
        string? ffmpeg = FindFfmpeg(Option(args, "--ffmpeg"));
        if (ffmpeg is null) {
            Console.Error.WriteLine("ffmpeg not found. It encodes the rendered frames: install it (winget install Gyan.FFmpeg) or pass --ffmpeg=<path>.");
            return 2;
        }
        Console.WriteLine("ffmpeg: " + ffmpeg);

        Directory.CreateDirectory(output);
        string scratch = Path.Combine(Environment.GetEnvironmentVariable("EVOTEC_SCRATCH_ROOT") is { Length: > 0 } scratchRoot ? scratchRoot : Path.Combine(output, ".frames"), "format-map-recorder");
        using var server = new StaticSite(site);
        int failures = 0;
        foreach (Cut cut in cuts) {
            try {
                await RecordAsync(server.BaseUrl, cut, output, Path.Combine(scratch, cut.Name), ffmpeg).ConfigureAwait(false);
            } catch (Exception ex) {
                failures++;
                Console.Error.WriteLine($"[{cut.Name}] failed: {ex.Message}");
                // A failed cut leaves no half-rendered frames behind.
                string leftover = Path.Combine(scratch, cut.Name);
                if (Directory.Exists(leftover)) Directory.Delete(leftover, recursive: true);
            }
        }
        return failures == 0 ? 0 : 1;
    }

    private static async Task RecordAsync(string baseUrl, Cut cut, string output, string frames, string ffmpeg) {
        var query = new List<string> { "ratio=" + cut.Ratio, "loop=0" };
        if (!string.IsNullOrWhiteSpace(cut.Scenes)) query.Add("scenes=" + Uri.EscapeDataString(cut.Scenes));
        if (!string.IsNullOrWhiteSpace(cut.Theme)) query.Add("theme=" + cut.Theme);
        if (Math.Abs(cut.Speed - 1) > 0.001) query.Add("speed=" + cut.Speed.ToString("0.##", CultureInfo.InvariantCulture));
        string url = $"{baseUrl}/format-map/?{string.Join('&', query)}";
        Console.WriteLine($"[{cut.Name}] rendering {cut.Width}x{cut.Height} at {cut.Fps} fps ({cut.Scale}x) {url}");

        if (Directory.Exists(frames)) Directory.Delete(frames, recursive: true);
        Directory.CreateDirectory(frames);

        // Width and Height are output pixels; the page lays out at Width/Scale CSS pixels and renders at Scale device pixels.
        var options = new HtmlBrowserLaunchOptions {
            Headless = true, Timeout = 120000, ViewportWidth = cut.Width / cut.Scale, ViewportHeight = cut.Height / cut.Scale, DeviceScaleFactor = cut.Scale
        };
        int count;
        var timer = Stopwatch.StartNew();
        await using (HtmlBrowserSession session = await HtmlBrowser.OpenSessionAsync("about:blank", options).ConfigureAwait(false)) {
            IPage page = session.Page;
            // Freeze time before the page's scripts exist; the page then only moves when RunFor says so.
            await page.Clock.InstallAsync().ConfigureAwait(false);
            await page.Clock.PauseAtAsync(DateTime.UtcNow.AddSeconds(5)).ConfigureAwait(false);
            await page.GotoAsync(url, new() { WaitUntil = WaitUntilState.Load }).ConfigureAwait(false);
            await page.EvaluateAsync("() => document.fonts.ready.then(() => true)").ConfigureAwait(false);
            string state = await page.EvaluateAsync<string>("() => document.querySelector('[data-format-map]')?.getAttribute('data-tour-state') ?? ''").ConfigureAwait(false);
            if (state != "playing") throw new InvalidOperationException($"The tour did not start (state '{state}'). Check the scene list: {cut.Scenes}");
            // A scene the page could not build (a renamed format, no route on that surface) means this is not the cut that cuts.json describes.
            string skipped = await page.EvaluateAsync<string>("() => document.querySelector('[data-format-map]').getAttribute('data-tour-skipped') ?? ''").ConfigureAwait(false);
            if (skipped.Length > 0) throw new InvalidOperationException($"The page has nothing to show for: {skipped.Replace("|", ", ")}. Update the scenes of cut '{cut.Name}' in Build/FormatMapRecorder/cuts.json.");
            double tourMs = double.Parse(await page.EvaluateAsync<string>("() => document.querySelector('[data-format-map]').getAttribute('data-tour-ms')").ConfigureAwait(false), CultureInfo.InvariantCulture);
            count = (int)Math.Ceiling((tourMs + 700) / 1000.0 * cut.Fps) + 1;

            long previous = 0;
            for (int index = 0; index < count; index++) {
                long at = (long)Math.Round(index * 1000.0 / cut.Fps);
                long delta = at - previous;
                if (delta > 0) await page.Clock.RunForAsync(delta).ConfigureAwait(false);
                await page.EvaluateAsync(StepScript, delta).ConfigureAwait(false);
                byte[] jpeg = await page.ScreenshotAsync(new() { Type = ScreenshotType.Jpeg, Quality = 95 }).ConfigureAwait(false);
                await File.WriteAllBytesAsync(Path.Combine(frames, $"{index:D6}.jpg"), jpeg).ConfigureAwait(false);
                previous = at;
                if (index % Math.Max(1, count / 10) == 0) Console.WriteLine($"  frame {index}/{count}  {timer.Elapsed.TotalSeconds:0}s");
            }
        }

        string[] input = ["-framerate", cut.Fps.ToString(CultureInfo.InvariantCulture), "-i", Path.Combine(frames, "%06d.jpg")];
        foreach (string format in cut.Formats.Select(static f => f.ToLowerInvariant()).Distinct()) {
            (string extension, string[] encode) = Encoding(format, cut);
            string target = Path.Combine(output, cut.Name + extension);
            await RunAsync(ffmpeg, ["-y", "-hide_banner", "-loglevel", "error", .. input, "-an", .. encode, target]).ConfigureAwait(false);
            Report(target);
        }
        string poster = Path.Combine(output, cut.Name + ".png");
        int posterFrame = Math.Clamp((int)(cut.PosterMs / 1000.0 * cut.Fps), 0, count - 1);
        await RunAsync(ffmpeg, ["-y", "-hide_banner", "-loglevel", "error", "-i", Path.Combine(frames, $"{posterFrame:D6}.jpg"), poster]).ConfigureAwait(false);
        Report(poster);
        Directory.Delete(frames, recursive: true);
        Console.WriteLine($"  rendered in {timer.Elapsed.TotalSeconds:0}s");
    }

    /// <summary>
    /// File name suffix and ffmpeg encoder arguments per output format.
    ///   mp4   H.264, plays everywhere (LinkedIn, X, Discord, YouTube)      h265  HEVC, smaller at the same quality
    ///   av1   AV1, smallest, newer players                                 webm  VP9
    ///   mov   ProRes 422 HQ 10-bit, an editing master (very large)         gif / webp / apng  silent animated loops
    /// </summary>
    private static (string Extension, string[] Arguments) Encoding(string format, Cut cut) {
        string loop = $"fps={cut.GifFps},scale={cut.GifWidth}:-1:flags=lanczos";
        string fps = cut.Fps.ToString(CultureInfo.InvariantCulture);
        return format switch {
            "mp4" => (".mp4", ["-c:v", "libx264", "-preset", "slow", "-crf", "17", "-pix_fmt", "yuv420p", "-r", fps, "-movflags", "+faststart"]),
            "h265" => (".h265.mp4", ["-c:v", "libx265", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p", "-tag:v", "hvc1", "-r", fps, "-movflags", "+faststart"]),
            "av1" => (".av1.mp4", ["-c:v", "libsvtav1", "-preset", "6", "-crf", "30", "-pix_fmt", "yuv420p", "-r", fps, "-movflags", "+faststart"]),
            "webm" => (".webm", ["-c:v", "libvpx-vp9", "-crf", "26", "-b:v", "0", "-r", fps]),
            "mov" => (".mov", ["-c:v", "prores_ks", "-profile:v", "3", "-vendor", "apl0", "-pix_fmt", "yuv422p10le", "-r", fps]),
            "gif" => (".gif", ["-vf", $"{loop},split[a][b];[a]palettegen=max_colors=128[p];[b][p]paletteuse=dither=bayer:bayer_scale=4", "-loop", "0"]),
            "webp" => (".webp", ["-vf", loop, "-c:v", "libwebp_anim", "-q:v", "75", "-loop", "0"]),
            "apng" => (".apng", ["-vf", loop, "-plays", "0"]),
            _ => throw new InvalidOperationException($"Unknown format '{format}' in cut '{cut.Name}'. Use mp4, h265, av1, webm, mov, gif, webp or apng.")
        };
    }

    private static void Report(string path) => Console.WriteLine($"  {Path.GetFileName(path),-38} {new FileInfo(path).Length / 1024.0 / 1024.0,7:0.00} MB");

    private static string? Option(string[] args, string name) =>
        args.FirstOrDefault(a => a.StartsWith(name + "=", StringComparison.OrdinalIgnoreCase))?[(name.Length + 1)..];

    private static string? FindFfmpeg(string? explicitPath) {
        if (!string.IsNullOrWhiteSpace(explicitPath)) return File.Exists(explicitPath) ? Path.GetFullPath(explicitPath) : null;
        string exe = OperatingSystem.IsWindows() ? "ffmpeg.exe" : "ffmpeg";
        foreach (string directory in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)) {
            string candidate = Path.Combine(directory.Trim('"'), exe);
            if (File.Exists(candidate)) return candidate;
        }
        string winget = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "WinGet", "Links", exe);
        return File.Exists(winget) ? winget : null;
    }

    private static async Task RunAsync(string executable, string[] arguments) {
        var info = new ProcessStartInfo(executable) { RedirectStandardError = true, RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
        foreach (string argument in arguments) info.ArgumentList.Add(argument);
        using Process process = Process.Start(info) ?? throw new InvalidOperationException("Could not start " + executable);
        Task<string> error = process.StandardError.ReadToEndAsync();
        Task<string> output = process.StandardOutput.ReadToEndAsync();
        await process.WaitForExitAsync().ConfigureAwait(false);
        await output.ConfigureAwait(false);
        if (process.ExitCode != 0) throw new InvalidOperationException($"ffmpeg exited with {process.ExitCode}: {(await error.ConfigureAwait(false)).Trim()}");
    }
}

internal sealed record CutFile(Cut[] Cuts);

/// <param name="Name">File name stem of the outputs.</param>
/// <param name="Ratio">16x9, 1x1, 4x5 or 9x16: the stage layout (data-ratio).</param>
/// <param name="Scenes">Scene selection (?scenes=); null plays the whole tour.</param>
/// <param name="Formats">mp4, h265, av1, webm, mov, gif, webp and/or apng.</param>
/// <param name="PosterMs">Where in the finished clip the poster frame is taken.</param>
/// <param name="Theme">dark or light; the stage defaults to dark.</param>
/// <param name="Scale">Device scale factor: Width and Height are output pixels, the layout uses Width/Scale CSS pixels.</param>
/// <param name="Fps">Frames per second of the render.</param>
internal sealed record Cut(string Name, string Ratio, int Width, int Height, string[] Formats, string? Scenes = null, double Speed = 1, int PosterMs = 2500, int GifWidth = 640, int GifFps = 12, string? Theme = null, int Scale = 1, int Fps = 30);

/// <summary>Serves a built website directory on a loopback port, so the render never needs a separate web server.</summary>
internal sealed class StaticSite : IDisposable {
    private static readonly Dictionary<string, string> ContentTypes = new(StringComparer.OrdinalIgnoreCase) {
        [".html"] = "text/html; charset=utf-8", [".css"] = "text/css; charset=utf-8", [".js"] = "text/javascript; charset=utf-8",
        [".json"] = "application/json", [".svg"] = "image/svg+xml", [".png"] = "image/png", [".jpg"] = "image/jpeg", [".jpeg"] = "image/jpeg",
        [".webp"] = "image/webp", [".woff2"] = "font/woff2", [".woff"] = "font/woff", [".ico"] = "image/x-icon", [".xml"] = "application/xml", [".txt"] = "text/plain; charset=utf-8"
    };

    private readonly HttpListener _listener = new();
    private readonly string _root;

    internal string BaseUrl { get; }

    internal StaticSite(string root) {
        _root = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        var probe = new TcpListener(IPAddress.Loopback, 0);
        probe.Start();
        int port = ((IPEndPoint)probe.LocalEndpoint).Port;
        probe.Stop();
        BaseUrl = $"http://127.0.0.1:{port}";
        _listener.Prefixes.Add(BaseUrl + "/");
        _listener.Start();
        _ = Task.Run(ServeAsync);
    }

    private async Task ServeAsync() {
        while (_listener.IsListening) {
            HttpListenerContext context;
            try {
                context = await _listener.GetContextAsync().ConfigureAwait(false);
            } catch (HttpListenerException) {
                return;
            } catch (ObjectDisposedException) {
                return;
            }
            _ = Task.Run(() => RespondAsync(context));
        }
    }

    private async Task RespondAsync(HttpListenerContext context) {
        try {
            string relative = Uri.UnescapeDataString(context.Request.Url?.AbsolutePath ?? "/").TrimStart('/');
            string file = Path.GetFullPath(Path.Combine(_root, relative));
            if (Directory.Exists(file)) file = Path.Combine(file, "index.html");
            if (!file.StartsWith(_root, StringComparison.OrdinalIgnoreCase) || !File.Exists(file)) {
                context.Response.StatusCode = 404;
            } else {
                byte[] bytes = await File.ReadAllBytesAsync(file).ConfigureAwait(false);
                context.Response.ContentType = ContentTypes.GetValueOrDefault(Path.GetExtension(file), "application/octet-stream");
                context.Response.ContentLength64 = bytes.Length;
                await context.Response.OutputStream.WriteAsync(bytes).ConfigureAwait(false);
            }
        } catch (Exception) {
            context.Response.StatusCode = 500;
        } finally {
            context.Response.Close();
        }
    }

    public void Dispose() => _listener.Close();
}
