using System.Buffers.Binary;
using System.Diagnostics;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed class ExternalPdfRasterizer {
    private ExternalPdfRasterizer(string executablePath, string textExtractorPath, string identity) {
        ExecutablePath = executablePath;
        TextExtractorPath = textExtractorPath;
        Identity = identity;
    }

    internal string ExecutablePath { get; }

    internal string TextExtractorPath { get; }

    internal string Identity { get; }

    internal static async Task<ExternalPdfRasterizer?> FindAsync(CancellationToken cancellationToken = default) {
        string? executablePath = FindOnPath(OperatingSystem.IsWindows() ? "pdftoppm.exe" : "pdftoppm");
        if (executablePath == null) return null;
        string? textExtractorPath = FindOnPath(OperatingSystem.IsWindows() ? "pdftotext.exe" : "pdftotext");
        if (textExtractorPath == null) return null;

        var startInfo = new ProcessStartInfo(executablePath) {
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add("-v");
        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start pdftoppm to read its version.");
        string standardOutput = await process.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        string standardError = await process.StandardError.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        string version = (standardError + Environment.NewLine + standardOutput)
            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault()?.Trim() ?? "pdftoppm";
        return new ExternalPdfRasterizer(executablePath, textExtractorPath, version);
    }

    internal async Task<HtmlPdfVisualEvidence> RenderFirstPageAsync(
        string pdfPath,
        string outputDirectory,
        string outputStem,
        CancellationToken cancellationToken = default) {
        string outputPrefix = Path.Combine(outputDirectory, outputStem);
        var startInfo = new ProcessStartInfo(ExecutablePath) {
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in new[] { "-f", "1", "-l", "1", "-singlefile", "-png", "-r", "120", pdfPath, outputPrefix }) {
            startInfo.ArgumentList.Add(argument);
        }

        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start pdftoppm for external visual evidence.");
        string standardOutput = await process.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        string standardError = await process.StandardError.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        if (process.ExitCode != 0) {
            throw new InvalidDataException(
                $"pdftoppm failed for '{Path.GetFileName(pdfPath)}' with exit code {process.ExitCode}: " +
                string.Join(" ", new[] { standardError, standardOutput }.Where(value => !string.IsNullOrWhiteSpace(value))));
        }

        string pngPath = outputPrefix + ".png";
        byte[] png = await File.ReadAllBytesAsync(pngPath, cancellationToken).ConfigureAwait(false);
        (int width, int height) = ReadPngDimensions(png);
        return new HtmlPdfVisualEvidence(
            Renderer: Identity,
            RelativePath: Path.GetFileName(pngPath),
            PageNumber: 1,
            Width: width,
            Height: height,
            SizeBytes: png.LongLength,
            Sha256: Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(png)).ToLowerInvariant(),
            Diagnostics: Array.Empty<string>());
    }

    internal async Task<IReadOnlyList<HtmlPdfVisualEvidence>> RenderAllPagesAsync(
        string pdfPath,
        string outputDirectory,
        string outputStem,
        double dpi,
        CancellationToken cancellationToken = default) {
        string outputPrefix = Path.Combine(outputDirectory, outputStem);
        var startInfo = new ProcessStartInfo(ExecutablePath) {
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in new[] {
            "-png",
            "-r", dpi.ToString(System.Globalization.CultureInfo.InvariantCulture),
            pdfPath,
            outputPrefix
        }) {
            startInfo.ArgumentList.Add(argument);
        }

        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start pdftoppm for external visual evidence.");
        string standardOutput = await process.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        string standardError = await process.StandardError.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        if (process.ExitCode != 0) {
            throw new InvalidDataException(
                $"pdftoppm failed for '{Path.GetFileName(pdfPath)}' with exit code {process.ExitCode}: " +
                string.Join(" ", new[] { standardError, standardOutput }.Where(value => !string.IsNullOrWhiteSpace(value))));
        }

        string filePrefix = outputStem + "-";
        string[] pagePaths = Directory.GetFiles(outputDirectory, filePrefix + "*.png")
            .Select(path => new { Path = path, Page = ParsePageNumber(path, filePrefix) })
            .Where(item => item.Page.HasValue)
            .OrderBy(item => item.Page)
            .Select(item => item.Path)
            .ToArray();
        if (pagePaths.Length == 0) {
            throw new InvalidDataException($"pdftoppm did not produce page images for '{Path.GetFileName(pdfPath)}'.");
        }

        var pages = new List<HtmlPdfVisualEvidence>(pagePaths.Length);
        foreach (string pagePath in pagePaths) {
            int pageNumber = ParsePageNumber(pagePath, filePrefix)
                ?? throw new InvalidDataException("Could not determine the pdftoppm page number.");
            byte[] png = await File.ReadAllBytesAsync(pagePath, cancellationToken).ConfigureAwait(false);
            (int width, int height) = ReadPngDimensions(png);
            pages.Add(new HtmlPdfVisualEvidence(
                Renderer: Identity,
                RelativePath: Path.GetFileName(pagePath),
                PageNumber: pageNumber,
                Width: width,
                Height: height,
                SizeBytes: png.LongLength,
                Sha256: Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(png)).ToLowerInvariant(),
                Diagnostics: Array.Empty<string>()));
        }
        return pages;
    }

    internal async Task<string> ExtractTextAsync(
        string pdfPath,
        CancellationToken cancellationToken = default) {
        var startInfo = new ProcessStartInfo(TextExtractorPath) {
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in new[] { "-layout", "-enc", "UTF-8", pdfPath, "-" }) {
            startInfo.ArgumentList.Add(argument);
        }

        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start pdftotext for external text evidence.");
        string standardOutput = await process.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        string standardError = await process.StandardError.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        if (process.ExitCode != 0) {
            throw new InvalidDataException(
                $"pdftotext failed for '{Path.GetFileName(pdfPath)}' with exit code {process.ExitCode}: " + standardError.Trim());
        }
        return standardOutput;
    }

    private static string? FindOnPath(string fileName) {
        string? path = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(path)) return null;
        foreach (string directory in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)) {
            string candidate = Path.Combine(directory.Trim(), fileName);
            if (File.Exists(candidate)) return Path.GetFullPath(candidate);
        }
        return null;
    }

    private static int? ParsePageNumber(string path, string filePrefix) {
        string fileName = Path.GetFileNameWithoutExtension(path);
        if (!fileName.StartsWith(filePrefix, StringComparison.Ordinal)) return null;
        return int.TryParse(
            fileName.AsSpan(filePrefix.Length),
            System.Globalization.NumberStyles.None,
            System.Globalization.CultureInfo.InvariantCulture,
            out int pageNumber)
            ? pageNumber
            : null;
    }

    private static (int Width, int Height) ReadPngDimensions(byte[] png) {
        if (png.Length < 24 || !png.AsSpan(0, 8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 })) {
            throw new InvalidDataException("pdftoppm did not produce a valid PNG preview.");
        }
        return (
            BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(16, 4)),
            BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(20, 4)));
    }
}
