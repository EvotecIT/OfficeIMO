using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.DjVu.Verification;

// Isolated opt-in comparison process. No reference decoder is linked, copied,
// downloaded, or invoked by an OfficeIMO runtime package.
internal static class DjVuReferenceRenderer {
    internal static async Task<string> GetVersionAsync(string executable) {
        using var process = Start(executable, "--help");
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        return ((await output) + (await error)).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "Unknown reference version";
    }

    internal static async Task<PageRenderingEvidence> CompareAsync(string executable, string input, int page, OfficeRasterImage image) {
        using var process = Start(executable, "-format=ppm", "-mode=color", "-page=" + page.ToString(CultureInfo.InvariantCulture), input);
        var error = process.StandardError.ReadToEndAsync();
        try {
            var stream = process.StandardOutput.BaseStream;
            string magic = Token(stream), widthToken = Token(stream), heightToken = Token(stream), maximum = Token(stream);
            int width = int.Parse(widthToken, CultureInfo.InvariantCulture), height = int.Parse(heightToken, CultureInfo.InvariantCulture);
            if (magic != "P6" || maximum != "255" || width != image.Width || height != image.Height)
                throw new InvalidDataException($"Reference raster header {magic} {width}x{height} {maximum} differs from managed {image.Width}x{image.Height}.");
            byte[] expected = new byte[checked(width * 3)], actual = new byte[expected.Length];
            long[] count = new long[3], sum = new long[3]; int[] max = new int[3];
            using var managedHash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            using var referenceHash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            for (int y = 0; y < height; y++) {
                await stream.ReadExactlyAsync(expected);
                for (int x = 0; x < width; x++) {
                    var color = image.GetPixel(x, y);
                    actual[x * 3] = color.R; actual[x * 3 + 1] = color.G; actual[x * 3 + 2] = color.B;
                }
                for (int i = 0; i < expected.Length; i++) {
                    int channel = i % 3, difference = Math.Abs(actual[i] - expected[i]);
                    if (difference != 0) count[channel]++;
                    sum[channel] += difference; max[channel] = Math.Max(max[channel], difference);
                }
                managedHash.AppendData(actual); referenceHash.AppendData(expected);
            }
            if (stream.ReadByte() != -1) throw new InvalidDataException("Unexpected data after the reference raster.");
            await process.WaitForExitAsync();
            if (process.ExitCode != 0) throw new IOException("Reference renderer failed: " + await error);
            return new PageRenderingEvidence { Width = width, Height = height, DifferingSamples = count, MaximumDifference = max,
                MeanDifference = sum.Select(v => v / ((double)width * height)).ToArray(),
                ManagedRgbSha256 = Convert.ToHexString(managedHash.GetHashAndReset()).ToLowerInvariant(),
                ReferenceRgbSha256 = Convert.ToHexString(referenceHash.GetHashAndReset()).ToLowerInvariant() };
        } finally {
            if (!process.HasExited) { process.Kill(entireProcessTree: true); await process.WaitForExitAsync(); }
            await error;
        }
    }

    private static Process Start(string executable, params string[] arguments) {
        var start = new ProcessStartInfo(executable) { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        return Process.Start(start) ?? throw new IOException("Could not start the explicitly selected reference renderer.");
    }

    private static string Token(Stream stream) {
        var token = new StringBuilder();
        int value;
        do {
            value = stream.ReadByte();
            if (value == '#') { do { value = stream.ReadByte(); } while (value != -1 && value != '\n'); }
            if (value == -1) throw new EndOfStreamException("Truncated reference PPM header.");
        } while (char.IsWhiteSpace((char)value));
        do {
            if (token.Length >= 32) throw new InvalidDataException("Excessive reference PPM header token.");
            token.Append((char)value); value = stream.ReadByte();
        } while (value != -1 && !char.IsWhiteSpace((char)value));
        if (value == -1) throw new EndOfStreamException("Truncated reference PPM header.");
        return token.ToString();
    }
}
