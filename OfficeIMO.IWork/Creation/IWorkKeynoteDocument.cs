using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>Creates native Keynote presentations containing positioned, uniformly styled text boxes.</summary>
/// <remarks>This creation model does not modify loaded iWork packages. Instances are not thread safe.</remarks>
public sealed class IWorkKeynoteDocument {
    private readonly List<IWorkKeynoteSlideBuilder> _slides = new();

    private IWorkKeynoteDocument(float widthPoints, float heightPoints) {
        IWorkKeynoteTextBox.ValidatePositive(widthPoints, nameof(widthPoints));
        IWorkKeynoteTextBox.ValidatePositive(heightPoints, nameof(heightPoints));
        SlideSize = new IWorkCanvasSize(widthPoints, heightPoints);
        Slides = _slides.AsReadOnly();
    }

    /// <summary>Gets the presentation canvas size in points.</summary>
    public IWorkCanvasSize SlideSize { get; }
    /// <summary>Gets the slides in presentation order.</summary>
    public IReadOnlyList<IWorkKeynoteSlideBuilder> Slides { get; }

    /// <summary>Creates an empty presentation with a 960 by 540 point canvas by default.</summary>
    public static IWorkKeynoteDocument Create(float widthPoints = 960, float heightPoints = 540) =>
        new(widthPoints, heightPoints);

    /// <summary>Adds a blank slide with an opaque six-digit sRGB background color.</summary>
    public IWorkKeynoteSlideBuilder AddSlide(string backgroundColor = "FFFFFF") {
        var slide = new IWorkKeynoteSlideBuilder(SlideSize, IWorkKeynoteTextBox.ParseColor(backgroundColor));
        _slides.Add(slide);
        return slide;
    }

    /// <summary>Encodes a deterministic native single-file Keynote package without a template or native application.</summary>
    /// <remarks>At least one slide is required. Bounds apply before output is returned; no font files are embedded.</remarks>
    public byte[] SaveBytes(IWorkKeynoteWriteOptions? options = null, CancellationToken cancellationToken = default) =>
        IWorkKeynotePackageWriter.Write(this, (options ?? new IWorkKeynoteWriteOptions()).Snapshot(), cancellationToken);

    /// <summary>Writes a completed package at the stream's current position and leaves the stream open.</summary>
    /// <remarks>Encoding finishes before the first write. Cancellation or an I/O failure during copying can leave partial stream data.</remarks>
    public void Save(Stream destination, IWorkKeynoteWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        byte[] bytes = SaveBytes(options, cancellationToken);
        Copy(bytes, destination, cancellationToken);
    }

    /// <summary>Atomically saves a native Keynote package, rejecting an existing path unless overwrite is enabled.</summary>
    /// <remarks>Validation and encoding finish before staging. Cancellation and failed staging preserve the existing destination.</remarks>
    public void Save(string path, bool overwrite = false, IWorkKeynoteWriteOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A destination path is required.", nameof(path));
        byte[] bytes = SaveBytes(options, cancellationToken);
        OfficeFileCommit.WriteAtomically(path, stream => Copy(bytes, stream, cancellationToken), cancellationToken,
            overwrite ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
    }

    private static void Copy(byte[] bytes, Stream destination, CancellationToken cancellationToken) {
        for (int offset = 0; offset < bytes.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(64 * 1024, bytes.Length - offset);
            destination.Write(bytes, offset, count);
            offset += count;
        }
        cancellationToken.ThrowIfCancellationRequested();
    }
}
