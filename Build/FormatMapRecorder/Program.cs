using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Sockets;
using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;

/// <summary>
/// Renders the format-map media diagram (Build/FormatMapMedia/index.html) into marketing clips: one clip per entry of its cuts.json,
/// framed for its platform (16:9, 1:1, 4:5, 9:16), plus a poster frame.
///
/// The diagram is a pure function of time (window.imoMedia.render(t)), so the render is frame by frame, not a screen recording.
/// A private loopback server serves the page, the generated Website/data/format_map.json and the site font; each frame renders one
/// moment and is screenshotted at the cut's device scale; ffmpeg encodes the frames. The result is the same on any machine at any
/// frame rate and resolution (--scale=2 renders 1920x1080 as true 3840x2160, sharp because the diagram is vector).
///
///   dotnet run --project Build/FormatMapRecorder -- [--output=Artefacts/FormatMapVideos] [--cut=name,name] [--scale=2] [--fps=60]
///                                                    [--formats=mp4,mov,...] [--still=2000,8000] [--cuts=file] [--media-dir=dir]
///                                                    [--ffmpeg=path] [--list]
/// Output formats: mp4 (H.264), h265, av1, webm (VP9), mov (ProRes 422 HQ), gif, webp, apng.
/// --still writes PNG stills of those moments (milliseconds) instead of video, for design work.
/// Frames are written under EVOTEC_SCRATCH_ROOT (or the output folder) and removed after a successful encode.
/// </summary>
internal static class Program {
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true, ReadCommentHandling = JsonCommentHandling.Skip };

    private static async Task<int> Main(string[] args) {
        string root = Directory.GetCurrentDirectory();
        string output = Path.GetFullPath(Option(args, "--output") ?? Path.Combine(root, "Artefacts", "FormatMapVideos"));
        string mediaDir = Path.GetFullPath(Option(args, "--media-dir") ?? Path.Combine(root, "Build", "FormatMapMedia"));
        string cutsPath = Option(args, "--cuts") ?? Path.Combine(mediaDir, "cuts.json");
        string dataDir = Path.Combine(root, "Website", "data");
        string fontDir = Path.Combine(root, "Website", "static", "fonts");
        if (!File.Exists(Path.Combine(mediaDir, "index.html")) || !File.Exists(Path.Combine(dataDir, "format_map.json")) || !File.Exists(cutsPath)) {
            Console.Error.WriteLine($"Needed: {mediaDir}/index.html, its cuts.json and Website/data/format_map.json. Run this from the repository root.");
            return 2;
        }
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
        double[] stills = (Option(args, "--still") ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(static s => double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double ms) ? ms : -1).Where(static ms => ms >= 0).ToArray();

        string? ffmpeg = FindFfmpeg(Option(args, "--ffmpeg"));
        if (ffmpeg is null) {
            Console.Error.WriteLine("ffmpeg not found. It encodes the rendered frames: install it (winget install Gyan.FFmpeg) or pass --ffmpeg=<path>.");
            return 2;
        }
        Console.WriteLine("ffmpeg: " + ffmpeg);

        Directory.CreateDirectory(output);
        string scratch = Path.Combine(Environment.GetEnvironmentVariable("EVOTEC_SCRATCH_ROOT") is { Length: > 0 } scratchRoot ? scratchRoot : Path.Combine(output, ".frames"), "format-map-recorder");
        using var server = new StaticSite(("/data/", dataDir), ("/fonts/", fontDir), ("/", mediaDir));
        int failures = 0;
        foreach (Cut cut in cuts) {
            try {
                await RecordAsync(server.BaseUrl, cut, output, Path.Combine(scratch, cut.Name), ffmpeg, stills).ConfigureAwait(false);
            } catch (Exception ex) {
                failures++;
                Console.Error.WriteLine($"[{cut.Name}] failed: {ex.Message}");
                // A failed cut leaves no half-rendered frames behind.
                string leftover = Path.Combine(scratch, cut.Name);
                if (Directory.Exists(leftover)) Directory.Delete(leftover, recursive: true);
            }
        }
        // Leave no empty scratch folder behind.
        if (Directory.Exists(scratch) && !Directory.EnumerateFileSystemEntries(scratch).Any()) Directory.Delete(scratch);
        return failures == 0 ? 0 : 1;
    }
    /// <summary>
    /// The media page (Build/FormatMapMedia/index.html) draws the diagram as a pure function of time, so nothing has to be frozen or stepped:
    /// each frame is imoMedia.render(t) followed by a screenshot.
    /// </summary>
    private static async Task RecordAsync(string baseUrl, Cut cut, string output, string frames, string ffmpeg, double[] stills) {
        var query = new List<string> { "ratio=" + cut.Ratio };
        if (!string.IsNullOrWhiteSpace(cut.Scenes)) query.Add("scenes=" + Uri.EscapeDataString(cut.Scenes));
        if (Math.Abs(cut.Speed - 1) > 0.001) query.Add("speed=" + cut.Speed.ToString("0.##", CultureInfo.InvariantCulture));
        string url = $"{baseUrl}/index.html?{string.Join('&', query)}";
        Console.WriteLine($"[{cut.Name}] rendering media {cut.Width}x{cut.Height} at {cut.Fps} fps ({cut.Scale}x) {url}");

        if (Directory.Exists(frames)) Directory.Delete(frames, recursive: true);
        Directory.CreateDirectory(frames);
        var options = new HtmlBrowserLaunchOptions {
            Headless = true, Timeout = 120000, ViewportWidth = cut.Width / cut.Scale, ViewportHeight = cut.Height / cut.Scale, DeviceScaleFactor = cut.Scale
        };
        int count;
        var timer = Stopwatch.StartNew();
        await using (HtmlBrowserSession session = await HtmlBrowser.OpenSessionAsync("about:blank", options).ConfigureAwait(false)) {
            IPage page = session.Page;
            var problems = new List<string>();
            page.PageError += (_, error) => problems.Add(error);
            await page.GotoAsync(url, new() { WaitUntil = WaitUntilState.Load }).ConfigureAwait(false);
            await page.EvaluateAsync("() => window.imoMedia.ready").ConfigureAwait(false);
            if (problems.Count > 0) throw new InvalidOperationException("The media page raised an error: " + problems[0]);
            string[] skipped = await page.EvaluateAsync<string[]>("() => window.imoMedia.skipped").ConfigureAwait(false);
            if (skipped.Length > 0) throw new InvalidOperationException($"The page has nothing to show for: {string.Join(", ", skipped)}. Update the scenes of cut '{cut.Name}' in Build/FormatMapMedia/cuts.json.");
            double duration = await page.EvaluateAsync<double>("() => window.imoMedia.duration").ConfigureAwait(false);
            if (stills.Length > 0) {
                foreach (double at in stills) {
                    await page.EvaluateAsync("(t) => window.imoMedia.render(t)", Math.Min(at, duration)).ConfigureAwait(false);
                    string still = Path.Combine(output, $"{cut.Name}-t{(int)at}.png");
                    await page.ScreenshotAsync(new() { Type = ScreenshotType.Png, Path = still }).ConfigureAwait(false);
                    Report(still);
                }
                Console.WriteLine($"  (duration {duration / 1000.0:0.0}s)");
                Directory.Delete(frames, recursive: true);
                return;
            }
            count = (int)Math.Ceiling(duration / 1000.0 * cut.Fps) + 1;
            for (int index = 0; index < count; index++) {
                await page.EvaluateAsync("(t) => window.imoMedia.render(t)", index * 1000.0 / cut.Fps).ConfigureAwait(false);
                byte[] jpeg = await page.ScreenshotAsync(new() { Type = ScreenshotType.Jpeg, Quality = 95 }).ConfigureAwait(false);
                await File.WriteAllBytesAsync(Path.Combine(frames, $"{index:D6}.jpg"), jpeg).ConfigureAwait(false);
                if (index % Math.Max(1, count / 10) == 0) Console.WriteLine($"  frame {index}/{count}  {timer.Elapsed.TotalSeconds:0}s");
            }
            if (problems.Count > 0) throw new InvalidOperationException("The media page raised an error while rendering: " + problems[0]);
        }
        await EncodeAsync(cut, output, frames, count, ffmpeg).ConfigureAwait(false);
        Console.WriteLine($"  rendered in {timer.Elapsed.TotalSeconds:0}s");
    }

    /// <summary>Encodes the numbered JPEG frames into every format of the cut, writes the poster frame and removes the frames.</summary>
    private static async Task EncodeAsync(Cut cut, string output, string frames, int count, string ffmpeg) {
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
/// <param name="Scale">Device scale factor: Width and Height are output pixels, the layout uses Width/Scale CSS pixels.</param>
/// <param name="Fps">Frames per second of the render.</param>
internal sealed record Cut(string Name, string Ratio, int Width, int Height, string[] Formats, string? Scenes = null, double Speed = 1, int PosterMs = 2500, int GifWidth = 640, int GifFps = 12, int Scale = 1, int Fps = 30);

/// <summary>Serves a built website directory on a loopback port, so the render never needs a separate web server.</summary>
internal sealed class StaticSite : IDisposable {
    private static readonly Dictionary<string, string> ContentTypes = new(StringComparer.OrdinalIgnoreCase) {
        [".html"] = "text/html; charset=utf-8", [".css"] = "text/css; charset=utf-8", [".js"] = "text/javascript; charset=utf-8",
        [".json"] = "application/json", [".svg"] = "image/svg+xml", [".png"] = "image/png", [".jpg"] = "image/jpeg", [".jpeg"] = "image/jpeg",
        [".webp"] = "image/webp", [".woff2"] = "font/woff2", [".woff"] = "font/woff", [".ico"] = "image/x-icon", [".xml"] = "application/xml", [".txt"] = "text/plain; charset=utf-8"
    };

    private readonly HttpListener _listener = new();
    private readonly (string Prefix, string Root)[] _mounts;

    internal string BaseUrl { get; }

    /// <summary>Serves each directory under its URL prefix; the longest matching prefix wins ("/" matches everything else).</summary>
    internal StaticSite(params (string Prefix, string Root)[] mounts) {
        _mounts = mounts
            .Select(static m => (m.Prefix, Root: Path.GetFullPath(m.Root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar))
            .OrderByDescending(static m => m.Prefix.Length)
            .ToArray();
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
            string requested = Uri.UnescapeDataString(context.Request.Url?.AbsolutePath ?? "/");
            (string prefix, string root) = _mounts.First(m => requested.StartsWith(m.Prefix, StringComparison.Ordinal));
            string file = Path.GetFullPath(Path.Combine(root, requested[prefix.Length..].TrimStart('/')));
            // The resolved path must stay inside the served folder before anything touches the file system.
            if (!file.StartsWith(root, StringComparison.OrdinalIgnoreCase)) {
                context.Response.StatusCode = 404;
                return;
            }
            if (Directory.Exists(file)) file = Path.Combine(file, "index.html");
            if (!File.Exists(file)) {
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
