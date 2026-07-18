using System;
using System.Diagnostics;
using System.IO;
using OfficeIMO.Pdf;
using OfficeIMO.Tests;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
    private static PdfDocument CreateVisualBaselineDocument(PdfOptions options) {
        return PdfDocument.Create(options);
    }

    private static void WriteReviewPdfArtifact(string scenarioName, byte[] pdfBytes) {
        string? outputDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT");
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        string safeName = scenarioName;
        foreach (char invalidChar in Path.GetInvalidFileNameChars()) {
            safeName = safeName.Replace(invalidChar, '-');
        }

        File.WriteAllBytes(Path.Combine(outputDirectory, safeName + ".pdf"), pdfBytes);
    }

    private static void AssertRasterBaseline(string baselineName, string actualPath) {
        string expectedPath = Path.Combine(GetPdfTestsProjectRoot(), "Pdf", "VisualBaselines", baselineName);
        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_PDF_RASTER_BASELINE"), "1", StringComparison.Ordinal)) {
            Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
            File.Copy(actualPath, expectedPath, overwrite: true);
            return;
        }

        if (!File.Exists(expectedPath)) {
            throw new FileNotFoundException(
                "PDF raster baseline missing. Set OFFICEIMO_UPDATE_PDF_RASTER_BASELINE=1 and re-run this test to generate it.",
                expectedPath);
        }

        VisualRasterComparison comparison = CompareRasterImages(File.ReadAllBytes(expectedPath), File.ReadAllBytes(actualPath));
        if (!comparison.Passed) {
            string artifactDirectory = VisualBaselineTestSupport.CreateArtifactDirectory("OfficeIMO.PdfRaster");

            string actualArtifactPath = Path.Combine(artifactDirectory, Path.GetFileName(actualPath));
            string expectedArtifactPath = Path.Combine(artifactDirectory, Path.GetFileName(expectedPath));
            string diffArtifactPath = Path.Combine(artifactDirectory, Path.GetFileNameWithoutExtension(actualPath) + ".diff.png");
            File.Copy(actualPath, actualArtifactPath, overwrite: true);
            File.Copy(expectedPath, expectedArtifactPath, overwrite: true);
            File.WriteAllBytes(diffArtifactPath, comparison.DiffPng);

            throw new Xunit.Sdk.XunitException(
                "PDF raster baseline changed. " +
                "Different pixels: " + comparison.DifferentPixels + "/" + comparison.TotalPixels + "; " +
                "max channel delta: " + comparison.MaxChannelDelta + "; " +
                "allowed different pixels: " + comparison.AllowedDifferentPixels + "; " +
                "channel tolerance: " + comparison.ChannelTolerance + ". " +
                "Artifacts: " + artifactDirectory + ".");
        }
    }

    private static VisualRasterComparison CompareRasterImages(byte[] expectedPng, byte[] actualPng) {
        int channelTolerance = VisualBaselineTestSupport.ReadNonNegativeInt("OFFICEIMO_PDF_RASTER_PIXEL_TOLERANCE", 0);
        int allowedDifferentPixels = VisualBaselineTestSupport.ReadNonNegativeInt("OFFICEIMO_PDF_RASTER_ALLOWED_DIFF_PIXELS", DefaultAllowedRasterNoisePixels);
        return CompareRasterImages(expectedPng, actualPng, channelTolerance, allowedDifferentPixels);
    }

    private static VisualRasterComparison CompareRasterImages(byte[] expectedPng, byte[] actualPng, int channelTolerance, int allowedDifferentPixels) =>
        VisualBaselineTestSupport.CompareRasterImages(expectedPng, actualPng, channelTolerance, allowedDifferentPixels);

    private static void RunPdftoppm(string rasterizerPath, string pdfPath, string outputPrefix, string workDir, int pageNumber) {
        string pageText = pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var psi = new ProcessStartInfo {
            FileName = rasterizerPath,
            Arguments = "-r 72 -png -singlefile -f " + pageText + " -l " + pageText + " " + Quote(pdfPath) + " " + Quote(outputPrefix),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = workDir
        };

        using var process = Process.Start(psi);
        if (process == null) {
            throw new InvalidOperationException("Could not start PDF rasterizer: " + rasterizerPath);
        }

        if (!process.WaitForExit(30000)) {
            process.Kill();
            throw new TimeoutException("PDF rasterizer timed out: " + rasterizerPath);
        }

        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        if (process.ExitCode != 0) {
            throw new InvalidOperationException("PDF rasterizer failed with exit code " + process.ExitCode + "." + Environment.NewLine + output + Environment.NewLine + error);
        }
    }

    private static bool TryFindPdftoppm(out string path) {
        string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_RASTERIZER");
        if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured)) {
            path = configured;
            return true;
        }

        string fileName = Environment.OSVersion.Platform == PlatformID.Win32NT ? "pdftoppm.exe" : "pdftoppm";
        string? pathVariable = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrWhiteSpace(pathVariable)) {
            foreach (string directory in pathVariable.Split(Path.PathSeparator)) {
                if (string.IsNullOrWhiteSpace(directory)) {
                    continue;
                }

                string candidate = Path.Combine(directory.Trim(), fileName);
                if (File.Exists(candidate)) {
                    path = candidate;
                    return true;
                }
            }
        }

        if (Environment.OSVersion.Platform == PlatformID.Win32NT) {
            string? localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");
            if (!string.IsNullOrWhiteSpace(localAppData)) {
                string packages = Path.Combine(localAppData, "Microsoft", "WinGet", "Packages");
                if (Directory.Exists(packages)) {
                    string[] candidates = Directory.GetFiles(packages, "pdftoppm.exe", SearchOption.AllDirectories);
                    if (candidates.Length > 0) {
                        path = candidates[0];
                        return true;
                    }
                }
            }
        }

        path = string.Empty;
        return false;
    }

    private static bool IsRequired() =>
        string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_PDF_RASTERIZER"), "1", StringComparison.Ordinal);

    private static bool CanAssertRasterBaseline(string rasterizerPath) {
        string currentVersion = ReadPdftoppmVersion(rasterizerPath);
        string versionPath = Path.Combine(GetPdfTestsProjectRoot(), "Pdf", "VisualBaselines", "poppler-version.txt");
        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_PDF_RASTER_BASELINE"), "1", StringComparison.Ordinal)) {
            Directory.CreateDirectory(Path.GetDirectoryName(versionPath)!);
            File.WriteAllText(versionPath, currentVersion + Environment.NewLine);
            return true;
        }

        string? baselineVersion = File.Exists(versionPath)
            ? File.ReadAllText(versionPath).Trim()
            : null;
        if (string.Equals(currentVersion, baselineVersion, StringComparison.Ordinal)) {
            return true;
        }

        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_PDF_RASTER_ALLOW_VERSION_MISMATCH"), "1", StringComparison.Ordinal)) {
            return true;
        }

        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_PDF_RASTER_BASELINE_MATCH"), "1", StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                "Strict PDF raster baselines were generated with Poppler " +
                (string.IsNullOrEmpty(baselineVersion) ? "(unrecorded)" : baselineVersion) +
                ", but the available pdftoppm is " + currentVersion + ". " +
                "Use the recorded Poppler version, deliberately update the baselines, or set " +
                "OFFICEIMO_PDF_RASTER_ALLOW_VERSION_MISMATCH=1 for an investigative comparison.");
        }

        return false;
    }

    private static string ReadPdftoppmVersion(string rasterizerPath) {
        var psi = new ProcessStartInfo {
            FileName = rasterizerPath,
            Arguments = "-v",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using Process process = Process.Start(psi)
            ?? throw new InvalidOperationException("Could not query PDF rasterizer version: " + rasterizerPath);
        string output = process.StandardOutput.ReadToEnd() + Environment.NewLine + process.StandardError.ReadToEnd();
        if (!process.WaitForExit(10000) || process.ExitCode != 0) {
            throw new InvalidOperationException("Could not query PDF rasterizer version: " + rasterizerPath);
        }

        System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(
            output,
            @"pdftoppm version (?<version>\d+\.\d+)",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant |
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!match.Success) {
            throw new InvalidOperationException("Could not parse pdftoppm version output: " + output.Trim());
        }

        return match.Groups["version"].Value;
    }

    private static bool SkipRasterAssertions() =>
        string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_SKIP_RASTER_ASSERTIONS"), "1", StringComparison.Ordinal);

    private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";

    private static string GetPdfTestsProjectRoot() =>
        VisualBaselineTestSupport.GetTestsProjectRoot();

    private static string GetTestsProjectRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "Images", "EvotecLogo.png"))) {
                return directory.FullName;
            }

            string aggregateProjectRoot = Path.Combine(directory.FullName, "OfficeIMO.Tests");
            if (File.Exists(Path.Combine(aggregateProjectRoot, "Images", "EvotecLogo.png"))) {
                return aggregateProjectRoot;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate OfficeIMO shared image fixtures from test runtime base directory.");
    }

    private static void TryDeleteDirectory(string directory) {
        try {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        } catch {
        }
    }
}
