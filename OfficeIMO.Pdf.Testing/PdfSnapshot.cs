using System.Globalization;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;

namespace OfficeIMO.Pdf.Testing;

/// <summary>
/// Verifies PDFs against approved snapshots: page images from the managed renderer, plus logical text, page sizes,
/// outline, form fields and render diagnostics. It works with any test framework because a mismatch throws
/// <see cref="PdfSnapshotMismatchException"/>.
/// </summary>
#pragma warning disable RS0026 // The required first parameter selects the PDF source; the caller-path default must stay last.
public static class PdfSnapshot {
    /// <summary>
    /// Environment variable that approves snapshots instead of comparing them when set to <c>1</c> or <c>true</c>.
    /// </summary>
    public const string AcceptEnvironmentVariable = "OFFICEIMO_PDF_SNAPSHOT_ACCEPT";

    private const string DefaultFolderName = "__pdfsnapshots__";

    /// <summary>Saves the document and verifies it against the approved snapshot <paramref name="name"/>.</summary>
    /// <exception cref="PdfSnapshotMismatchException">The PDF differs from the approved snapshot, or none was approved yet.</exception>
    public static void Verify(PdfDocument document, string name, PdfSnapshotOptions? options = null, [CallerFilePath] string callerFilePath = "") {
        ArgumentNullException.ThrowIfNull(document);
        Verify(document.ToBytes(), name, options, callerFilePath);
    }

    /// <summary>Reads the stream from its current position and verifies it against the approved snapshot <paramref name="name"/>.</summary>
    /// <exception cref="PdfSnapshotMismatchException">The PDF differs from the approved snapshot, or none was approved yet.</exception>
    public static void Verify(Stream pdf, string name, PdfSnapshotOptions? options = null, [CallerFilePath] string callerFilePath = "") {
        ArgumentNullException.ThrowIfNull(pdf);
        using var buffer = new MemoryStream();
        pdf.CopyTo(buffer);
        Verify(buffer.ToArray(), name, options, callerFilePath);
    }

    /// <summary>Verifies PDF bytes against the approved snapshot <paramref name="name"/>.</summary>
    /// <remarks>
    /// Checks run in order: identical PDF bytes when <see cref="PdfSnapshotOptions.Pdf"/> stores them, then the structure
    /// file (exact), then page pixels within <see cref="PdfSnapshotOptions.Visual"/> tolerances. On a mismatch the received
    /// files and an HTML review gallery are written next to the approved files.
    /// </remarks>
    /// <exception cref="PdfSnapshotMismatchException">The PDF differs from the approved snapshot, or none was approved yet.</exception>
    public static void Verify(byte[] pdf, string name, PdfSnapshotOptions? options = null, [CallerFilePath] string callerFilePath = "") {
        ArgumentNullException.ThrowIfNull(pdf);
        PdfSnapshotFiles.ValidateName(name);
        PdfSnapshotOptions effective = options ?? new PdfSnapshotOptions();
        var files = new PdfSnapshotFiles(ResolveDirectory(effective, callerFilePath), name);
        byte[]? storedPdf = effective.Pdf == PdfSnapshotPdfMode.None ? null : pdf;

        if (storedPdf is not null && !IsAccepting() && File.Exists(files.Pdf(received: false))
            && File.Exists(files.Structure(received: false)) && File.ReadAllBytes(files.Pdf(received: false)).SequenceEqual(pdf)) {
            files.DeleteReceived();
            return;
        }

        PdfSnapshotCapture capture = PdfSnapshotCapture.Create(pdf, name, effective);
        if (IsAccepting()) {
            files.WriteApproved(capture, storedPdf);
            return;
        }

        if (!File.Exists(files.Structure(received: false))) {
            List<string> created = files.WriteReceived(capture, storedPdf);
            throw new PdfSnapshotMismatchException(name,
                "PDF snapshot '" + name + "' has no approved baseline in " + files.Directory + ".\n"
                + "Review the received files, then approve them:\n" + AcceptHint(files.Directory) + ReceivedList(created),
                created.AsReadOnly(), reviewPath: null);
        }

        var problems = new StringBuilder();
        if (effective.Pdf == PdfSnapshotPdfMode.Exact) CompareBytes(files, pdf, problems);
        string approvedStructure = File.ReadAllText(files.Structure(received: false));
        if (PdfSnapshotDiff.Normalize(approvedStructure) != capture.Structure) {
            problems.Append("Structure differs (").Append(Path.GetFileName(files.Structure(received: false))).Append("):\n")
                .Append(PdfSnapshotDiff.Describe(approvedStructure, capture.Structure));
            if (PdfSnapshotCapture.FontSection(approvedStructure) != PdfSnapshotCapture.FontSection(capture.Structure)) {
                problems.Append("Embedded fonts differ from the approved snapshot. A PDF that embeds fonts installed on the machine, ")
                    .Append("as OfficeIMO.Pdf does by default, renders differently elsewhere; register font files with ")
                    .Append("PdfCreateOptions.RegisterFontFamily so every machine produces the same PDF.\n");
            }
        }
        string[] missingPages = Enumerable.Range(1, PdfSnapshotCapture.PageCount(approvedStructure))
            .Select(page => files.Page(received: false, page)).Where(path => !File.Exists(path)).Select(path => Path.GetFileName(path)!).ToArray();
        if (missingPages.Length > 0) {
            problems.Append("Approved page images are missing: ").Append(string.Join(", ", missingPages))
                .Append(". Restore them from version control or approve the snapshot again.\n");
        }
        PdfVisualComparisonReport visual = PdfVisualComparer.CompareImages(files.ReadApprovedPages(), capture.Pages, effective.Visual);
        DescribeVisual(visual, problems);
        if (problems.Length == 0) {
            files.DeleteReceived();
            return;
        }

        List<string> received = files.WriteReceived(capture, storedPdf);
        File.WriteAllText(files.Review, visual.ToHtmlGallery("PDF snapshot " + name), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        received.Add(files.Review);
        throw new PdfSnapshotMismatchException(name,
            "PDF snapshot '" + name + "' does not match the approved baseline.\n" + problems
            + "Review: " + files.Review + "\n" + ReceivedList(received)
            + "If the change is intended, approve it:\n" + AcceptHint(files.Directory),
            received.AsReadOnly(), files.Review);
    }

    /// <summary>Approves every snapshot in <paramref name="directory"/> that has received files waiting for review.</summary>
    /// <returns>The number of snapshots approved.</returns>
    public static int Accept(string directory) {
        if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("A snapshot directory is required.", nameof(directory));
        if (!Directory.Exists(directory)) return 0;
        int accepted = 0;
        foreach (string name in PdfSnapshotFiles.PendingNames(directory)) {
            new PdfSnapshotFiles(directory, name).Accept();
            accepted++;
        }
        return accepted;
    }

    private static bool IsAccepting() {
        string? value = Environment.GetEnvironmentVariable(AcceptEnvironmentVariable);
        return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolveDirectory(PdfSnapshotOptions options, string callerFilePath) {
        if (!string.IsNullOrWhiteSpace(options.Directory)) return Path.GetFullPath(options.Directory!);
        string? sourceDirectory = string.IsNullOrEmpty(callerFilePath) ? null : Path.GetDirectoryName(callerFilePath);
        if (sourceDirectory is null || !Directory.Exists(sourceDirectory)) {
            throw new InvalidOperationException(
                "PDF snapshots are stored next to the calling test file, but its folder '" + (sourceDirectory ?? "(unknown)")
                + "' does not exist on this machine. Deterministic or path-mapped builds hide it; set PdfSnapshotOptions.Directory.");
        }
        return Path.Combine(sourceDirectory, DefaultFolderName);
    }

    private static void CompareBytes(PdfSnapshotFiles files, byte[] pdf, StringBuilder problems) {
        string approvedPath = files.Pdf(received: false);
        string approved = File.Exists(approvedPath) ? Sha256(File.ReadAllBytes(approvedPath)) : "(no approved PDF)";
        string actual = Sha256(pdf);
        if (approved != actual) {
            problems.Append("PDF bytes differ: approved SHA-256 ").Append(approved).Append(", received ").Append(actual).Append(".\n");
        }
    }

    private static void DescribeVisual(PdfVisualComparisonReport report, StringBuilder problems) {
        foreach (string difference in report.StructuralDifferences) problems.Append("Pages: ").Append(difference).Append('\n');
        foreach (PdfVisualPageComparison page in report.Pages.Where(static page => !page.IsMatch)) {
            problems.Append("Page ").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture)).Append(" pixels differ: ")
                .Append(page.DifferentPixels.ToString("N0", CultureInfo.InvariantCulture)).Append(" changed (")
                .Append((page.DifferenceRatio * 100D).ToString("0.###", CultureInfo.InvariantCulture)).Append("%)");
            if (page.ChangedBounds is PdfPixelRegion bounds) {
                problems.Append(" within x=").Append(bounds.X).Append(" y=").Append(bounds.Y)
                    .Append(" w=").Append(bounds.Width).Append(" h=").Append(bounds.Height);
            }
            if (page.HasSizeDifference) problems.Append("; page size changed");
            problems.Append(".\n");
        }
    }

    private static string ReceivedList(IEnumerable<string> paths) =>
        "Received:\n" + string.Concat(paths.Select(static path => "  " + path + "\n"));

    private static string AcceptHint(string directory) =>
        "  set " + AcceptEnvironmentVariable + "=1 and run the tests again, or call PdfSnapshot.Accept(\"" + directory + "\").\n";

    private static string Sha256(byte[] bytes) {
#if NET8_0_OR_GREATER
        byte[] hash = SHA256.HashData(bytes);
#else
        using SHA256 sha = SHA256.Create();
        byte[] hash = sha.ComputeHash(bytes);
#endif
        return string.Concat(hash.Select(static value => value.ToString("x2", CultureInfo.InvariantCulture)));
    }
}
#pragma warning restore RS0026
