using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Pdf.Testing;

/// <summary>
/// File layout of one snapshot. Approved files are <c>name.structure.json</c>, <c>name.page-001.png</c> and optionally
/// <c>name.pdf</c>; received files insert <c>.received</c> after the name, and the review gallery is <c>name.received.html</c>.
/// </summary>
internal sealed class PdfSnapshotFiles {
    internal const string ReceivedMarker = ".received";
    private const string StructureSuffix = ".structure.json";
    private const string PageMarker = ".page-";
    private static readonly char[] PathSeparators = { '/', '\\', ':' };

    public PdfSnapshotFiles(string directory, string name) {
        Directory = directory;
        Name = name;
    }

    public string Directory { get; }
    public string Name { get; }

    public string Structure(bool received) => Path.Combine(Directory, Stem(received) + StructureSuffix);
    public string Pdf(bool received) => Path.Combine(Directory, Stem(received) + ".pdf");
    public string Page(bool received, int pageNumber) =>
        Path.Combine(Directory, Stem(received) + PageMarker + pageNumber.ToString("000", CultureInfo.InvariantCulture) + ".png");
    public string Review => Path.Combine(Directory, Name + ReceivedMarker + ".html");

    public static void ValidateName(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Snapshot name cannot be empty.", nameof(name));
        if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || name.IndexOfAny(PathSeparators) >= 0) {
            throw new ArgumentException("Snapshot name '" + name + "' must be a plain file name.", nameof(name));
        }
        if (ContainsIgnoreCase(name, ReceivedMarker) || ContainsIgnoreCase(name, PageMarker)) {
            throw new ArgumentException("Snapshot name '" + name + "' cannot contain '" + ReceivedMarker + "' or '" + PageMarker + "'.", nameof(name));
        }
    }

#pragma warning disable CA2249 // string.Contains(string, StringComparison) is unavailable on netstandard2.0 and net472.
    private static bool ContainsIgnoreCase(string value, string part) => value.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0;
#pragma warning restore CA2249

    /// <summary>Names of snapshots in a folder that have received files waiting for review.</summary>
    public static IEnumerable<string> PendingNames(string directory) {
        string suffix = ReceivedMarker + StructureSuffix;
        return System.IO.Directory.EnumerateFiles(directory, "*" + suffix)
            .Select(Path.GetFileName)
            .Select(file => file!.Substring(0, file.Length - suffix.Length))
            .OrderBy(static name => name, StringComparer.Ordinal)
            .ToArray();
    }

    public IReadOnlyList<byte[]> ReadApprovedPages() {
        var pages = new List<byte[]>();
        for (int page = 1; File.Exists(Page(received: false, page)); page++) pages.Add(File.ReadAllBytes(Page(received: false, page)));
        return pages;
    }

    /// <summary>Writes a capture as received files and returns their paths in a stable order.</summary>
    public List<string> WriteReceived(PdfSnapshotCapture capture, byte[]? pdf) {
        DeleteSet(received: true);
        return Write(received: true, capture, pdf);
    }

    public void WriteApproved(PdfSnapshotCapture capture, byte[]? pdf) {
        DeleteSet(received: false);
        Write(received: false, capture, pdf);
        DeleteReceived();
    }

    /// <summary>Promotes received files to approved files.</summary>
    public void Accept() {
        // Keep the approved baseline unless there is a received snapshot to replace it.
        if (!File.Exists(Structure(received: true))) throw new FileNotFoundException("No received snapshot to accept.", Structure(received: true));
        DeleteSet(received: false);
        File.Move(Structure(received: true), Structure(received: false));
        if (File.Exists(Pdf(received: true))) File.Move(Pdf(received: true), Pdf(received: false));
        for (int page = 1; File.Exists(Page(received: true, page)); page++) File.Move(Page(received: true, page), Page(received: false, page));
        DeleteReceived();
    }

    public void DeleteReceived() {
        DeleteSet(received: true);
        if (File.Exists(Review)) File.Delete(Review);
    }

    private List<string> Write(bool received, PdfSnapshotCapture capture, byte[]? pdf) {
        System.IO.Directory.CreateDirectory(Directory);
        var written = new List<string> { Structure(received) };
        File.WriteAllText(Structure(received), capture.Structure, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        for (int index = 0; index < capture.Pages.Count; index++) {
            string path = Page(received, index + 1);
            File.WriteAllBytes(path, Compact(capture.Pages[index]));
            written.Add(path);
        }
        if (pdf is not null) {
            File.WriteAllBytes(Pdf(received), pdf);
            written.Add(Pdf(received));
        }
        return written;
    }

    private void DeleteSet(bool received) {
        if (!System.IO.Directory.Exists(Directory)) return;
        if (File.Exists(Structure(received))) File.Delete(Structure(received));
        if (File.Exists(Pdf(received))) File.Delete(Pdf(received));
        for (int page = 1; File.Exists(Page(received, page)); page++) File.Delete(Page(received, page));
    }

    /// <summary>Re-encodes a page image with the smallest PNG settings when it is written, since approved images are committed.</summary>
    private static byte[] Compact(byte[] png) {
        if (!OfficePngReader.TryDecode(png, out OfficeRasterImage? image) || image is null) return png;
        byte[] smallest = OfficePngWriter.Encode(image, new OfficePngEncodeOptions { Effort = OfficeImageEncodingEffort.Smallest });
        return smallest.Length < png.Length ? smallest : png;
    }

    private string Stem(bool received) => received ? Name + ReceivedMarker : Name;
}
