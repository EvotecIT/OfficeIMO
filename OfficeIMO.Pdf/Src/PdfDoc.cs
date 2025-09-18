using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Root document container and fluent API for composing simple PDFs.
/// Mirrors the OfficeIMO.Markdown style with H1/H2/H3 and P blocks.
/// </summary>
public sealed class PdfDoc {
    private readonly List<IPdfBlock> _blocks = new();
    private readonly PdfOptions _options;

    // Metadata
    private string? _title;
    private string? _author;
    private string? _subject;
    private string? _keywords;

    private PdfDoc(PdfOptions? options = null) { _options = options ?? new PdfOptions(); }

    public static PdfDoc Create(PdfOptions? options = null) => new PdfDoc(options);

    public PdfDoc Meta(string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        _title = title ?? _title;
        _author = author ?? _author;
        _subject = subject ?? _subject;
        _keywords = keywords ?? _keywords;
        return this;
    }

    public PdfDoc H1(string text) { _blocks.Add(new HeadingBlock(1, text)); return this; }
    public PdfDoc H2(string text) { _blocks.Add(new HeadingBlock(2, text)); return this; }
    public PdfDoc H3(string text) { _blocks.Add(new HeadingBlock(3, text)); return this; }
    public PdfDoc P(string text) { _blocks.Add(new ParagraphBlock(text)); return this; }
    public PdfDoc PageBreak() { _blocks.Add(new PageBreakBlock()); return this; }

    public byte[] ToBytes() => PdfWriter.Write(this, _blocks, _options, _title, _author, _subject, _keywords);

    public PdfDoc Save(string path) {
        var bytes = ToBytes();
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)) ?? ".");
        File.WriteAllBytes(path, bytes);
        return this;
    }

    public async System.Threading.Tasks.Task SaveAsync(string path) {
        var bytes = ToBytes();
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)) ?? ".");
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
        await fs.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
    }

    // Internal getters for writer
    internal IEnumerable<IPdfBlock> Blocks => _blocks;
    internal PdfOptions Options => _options;
}

