using System.Diagnostics;
using System.Globalization;

/// <summary>Validates the selected FFmpeg encoders and stages complete cut artifacts before replacing existing output.</summary>
internal static class MediaEncoding {
    internal static async Task ValidateAsync(string ffmpeg, Cut[] cuts) {
        string listing = await RunAsync(ffmpeg, ["-hide_banner", "-encoders"]).ConfigureAwait(false);
        var available = listing.Split('\n').Select(line => line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
            .Where(parts => parts.Length >= 2 && parts[0].Length == 6 && parts[0][0] == 'V')
            .Select(parts => parts[1]).ToHashSet(StringComparer.Ordinal);
        var required = cuts.SelectMany(cut => cut.Formats.Select(format => {
            string[] arguments = Settings(format.ToLowerInvariant(), cut).Arguments;
            return arguments[Array.IndexOf(arguments, "-c:v") + 1];
        })).Append("png").Distinct(StringComparer.Ordinal);
        string[] missing = required.Where(encoder => !available.Contains(encoder)).ToArray();
        if (missing.Length > 0) throw new ArgumentException("FFmpeg is missing required video encoders: " + string.Join(", ", missing));
    }

    internal static async Task EncodeAsync(Cut cut, string output, string frames, string ffmpeg) {
        var formats = cut.Formats.Select(f => Settings(f.ToLowerInvariant(), cut)).DistinctBy(f => f.Extension).ToArray();
        string poster = cut.Name + ".png";
        string[] names = formats.Select(f => cut.Name + f.Extension).Append(poster).ToArray();
        await MediaArtifacts.PublishAsync(output, names, async staging => {
            string[] input = ["-framerate", cut.Fps.ToString(CultureInfo.InvariantCulture), "-i", Path.Combine(frames, "%06d.jpg")];
            foreach ((string extension, string[] encode) in formats) {
                string name = cut.Name + extension;
                await RunAsync(ffmpeg, ["-y", "-hide_banner", "-loglevel", "error", .. input, "-an", .. encode, Path.Combine(staging, name)]).ConfigureAwait(false);
            }
            int posterFrame = (int)(cut.PosterMs / 1000.0 * cut.Fps);
            await RunAsync(ffmpeg, ["-y", "-hide_banner", "-loglevel", "error", "-i", Path.Combine(frames, $"{posterFrame:D6}.jpg"), "-c:v", "png", Path.Combine(staging, poster)]).ConfigureAwait(false);
        }).ConfigureAwait(false);
        Directory.Delete(frames, recursive: true);
    }

    /// <summary>File extension and explicit FFmpeg encoder arguments for each supported format.</summary>
    internal static (string Extension, string[] Arguments) Settings(string format, Cut cut) {
        string loop = $"fps={cut.GifFps},scale={cut.GifWidth}:-1:flags=lanczos";
        string fps = cut.Fps.ToString(CultureInfo.InvariantCulture);
        return format switch {
            "mp4" => (".mp4", ["-c:v", "libx264", "-preset", "slow", "-crf", "17", "-pix_fmt", "yuv420p", "-r", fps, "-movflags", "+faststart"]),
            "h265" => (".h265.mp4", ["-c:v", "libx265", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p", "-tag:v", "hvc1", "-r", fps, "-movflags", "+faststart"]),
            "av1" => (".av1.mp4", ["-c:v", "libsvtav1", "-preset", "6", "-crf", "30", "-pix_fmt", "yuv420p", "-r", fps, "-movflags", "+faststart"]),
            "webm" => (".webm", ["-c:v", "libvpx-vp9", "-crf", "26", "-b:v", "0", "-r", fps]),
            "mov" => (".mov", ["-c:v", "prores_ks", "-profile:v", "3", "-vendor", "apl0", "-pix_fmt", "yuv422p10le", "-r", fps]),
            "gif" => (".gif", ["-c:v", "gif", "-vf", $"{loop},split[a][b];[a]palettegen=max_colors=128[p];[b][p]paletteuse=dither=bayer:bayer_scale=4", "-loop", "0"]),
            "webp" => (".webp", ["-c:v", "libwebp_anim", "-vf", loop, "-q:v", "75", "-loop", "0"]),
            "apng" => (".apng", ["-c:v", "apng", "-vf", loop, "-plays", "0"]),
            _ => throw new InvalidOperationException($"Unknown format '{format}' in cut '{cut.Name}'. Use mp4, h265, av1, webm, mov, gif, webp or apng.")
        };
    }

    private static async Task<string> RunAsync(string executable, string[] arguments) {
        var info = new ProcessStartInfo(executable) { RedirectStandardError = true, RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
        foreach (string argument in arguments) info.ArgumentList.Add(argument);
        using Process process = Process.Start(info) ?? throw new InvalidOperationException("Could not start " + executable);
        Task<string> error = process.StandardError.ReadToEndAsync();
        Task<string> output = process.StandardOutput.ReadToEndAsync();
        await process.WaitForExitAsync().ConfigureAwait(false);
        string result = await output.ConfigureAwait(false);
        string diagnostic = (await error.ConfigureAwait(false)).Trim();
        if (process.ExitCode != 0) throw new InvalidOperationException($"ffmpeg exited with {process.ExitCode}: {diagnostic}");
        return result;
    }
}
