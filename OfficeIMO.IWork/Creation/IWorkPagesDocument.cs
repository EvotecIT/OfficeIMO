using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>Creates template-free native Pages body documents with page settings and rich paragraphs.</summary>
/// <remarks>Instances are not thread safe. This model does not modify loaded packages or measure pagination.</remarks>
public sealed class IWorkPagesDocument {
    private readonly List<IWorkPagesParagraphData> _paragraphs = new();

    private IWorkPagesDocument(float width, float height, float left, float right, float top, float bottom) {
        IWorkCreationValidation.ValidatePositive(width, nameof(width));
        IWorkCreationValidation.ValidatePositive(height, nameof(height));
        foreach (var margin in new[] { (left, nameof(left)), (right, nameof(right)), (top, nameof(top)), (bottom, nameof(bottom)) }) {
            IWorkCreationValidation.ValidateFinite(margin.Item1, margin.Item2);
            if (margin.Item1 < 0) throw new ArgumentOutOfRangeException(margin.Item2, "Margins cannot be negative.");
        }
        if ((double)left + right >= width || (double)top + bottom >= height)
            throw new ArgumentException("Margins must leave a positive body area.");
        PageSize = new IWorkCanvasSize(width, height);
        LeftMarginPoints = left; RightMarginPoints = right; TopMarginPoints = top; BottomMarginPoints = bottom;
        Paragraphs = _paragraphs.AsReadOnly();
    }

    /// <summary>Gets page dimensions in points.</summary>
    public IWorkCanvasSize PageSize { get; }
    /// <summary>Gets the left body margin in points.</summary>
    public float LeftMarginPoints { get; }
    /// <summary>Gets the right body margin in points.</summary>
    public float RightMarginPoints { get; }
    /// <summary>Gets the top body margin in points.</summary>
    public float TopMarginPoints { get; }
    /// <summary>Gets the bottom body margin in points.</summary>
    public float BottomMarginPoints { get; }
    /// <summary>Gets the body paragraphs in insertion order.</summary>
    public IReadOnlyList<IWorkPagesParagraphData> Paragraphs { get; }

    /// <summary>Creates an empty document, using US Letter and one-inch margins by default.</summary>
    public static IWorkPagesDocument Create(float widthPoints = 612, float heightPoints = 792,
        float leftMarginPoints = 72, float rightMarginPoints = 72, float topMarginPoints = 72, float bottomMarginPoints = 72) =>
        new(widthPoints, heightPoints, leftMarginPoints, rightMarginPoints, topMarginPoints, bottomMarginPoints);

    /// <summary>Adds a left-aligned body paragraph with one font and opaque sRGB color.</summary>
    /// <remarks>Use a separate call for each paragraph. Empty text creates an editable blank paragraph.</remarks>
    public IWorkPagesParagraphData AddParagraph(string text, string fontName = "Arial", float fontSizePoints = 12,
        string color = "000000") => AddRichParagraph(new[] { new IWorkTextRunData(text, fontName, fontSizePoints, color) });

    /// <summary>Adds a snapshot of rich character runs as one left-aligned body paragraph.</summary>
    /// <remarks>Runs cannot contain line breaks. Empty ranges are discarded unless the whole paragraph is blank.</remarks>
    public IWorkPagesParagraphData AddRichParagraph(IEnumerable<IWorkTextRunData> runs) {
        var paragraph = new IWorkPagesParagraphData(runs);
        _paragraphs.Add(paragraph);
        return paragraph;
    }

    /// <summary>Encodes a deterministic native Pages package without a template or native application.</summary>
    /// <remarks>At least one paragraph is required. No font files are embedded; native fonts and pagination apply.</remarks>
    public byte[] SaveBytes(IWorkPagesWriteOptions? options = null, CancellationToken cancellationToken = default) =>
        IWorkPagesPackageWriter.Write(this, (options ?? new IWorkPagesWriteOptions()).Snapshot(), cancellationToken);

    /// <summary>Writes a completed package at the stream's current position and leaves the stream open.</summary>
    /// <remarks>Encoding finishes before writing; cancellation or I/O failure while copying can leave partial stream data.</remarks>
    public void Save(Stream destination, IWorkPagesWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        IWorkNativePackageWriter.Copy(SaveBytes(options, cancellationToken), destination, cancellationToken);
    }

    /// <summary>Atomically saves a native package, rejecting an existing file unless overwrite is enabled.</summary>
    /// <remarks>Validation and encoding finish before staging. Cancellation or failed staging preserves the existing destination.</remarks>
    public void Save(string path, bool overwrite = false, IWorkPagesWriteOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A destination path is required.", nameof(path));
        byte[] bytes = SaveBytes(options, cancellationToken);
        OfficeFileCommit.WriteAtomically(path, stream => IWorkNativePackageWriter.Copy(bytes, stream, cancellationToken),
            cancellationToken, overwrite ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
    }
}
