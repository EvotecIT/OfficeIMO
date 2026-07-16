namespace OfficeIMO.OneNote;

/// <summary>
/// Lazily materialized binary data associated with a OneNote image, file, recording, or ink object.
/// </summary>
public sealed class OneNoteBinaryPayload {
    private readonly byte[]? _bytes;
    private readonly Func<Stream>? _streamFactory;

    private OneNoteBinaryPayload(byte[]? bytes, Func<Stream>? streamFactory, long? length) {
        _bytes = bytes;
        _streamFactory = streamFactory;
        Length = length;
    }

    /// <summary>Payload length when known without opening the stream.</summary>
    public long? Length { get; }

    /// <summary>Creates an independently owned payload by copying a byte array.</summary>
    public static OneNoteBinaryPayload FromBytes(byte[] bytes) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        var copy = new byte[bytes.Length];
        Buffer.BlockCopy(bytes, 0, copy, 0, bytes.Length);
        return new OneNoteBinaryPayload(copy, null, copy.LongLength);
    }

    /// <summary>
    /// Creates a lazy payload. The factory must return a new readable stream for every call.
    /// </summary>
    public static OneNoteBinaryPayload FromStreamFactory(Func<Stream> streamFactory, long? length = null) {
        if (streamFactory == null) throw new ArgumentNullException(nameof(streamFactory));
        if (length.HasValue && length.Value < 0) throw new ArgumentOutOfRangeException(nameof(length));
        return new OneNoteBinaryPayload(null, streamFactory, length);
    }

    /// <summary>Opens a new readable stream for the payload.</summary>
    public Stream OpenRead() {
        if (_bytes != null) {
            return new MemoryStream(_bytes, 0, _bytes.Length, false, true);
        }

        Stream stream = _streamFactory!();
        if (stream == null) throw new InvalidOperationException("The OneNote payload stream factory returned null.");
        if (!stream.CanRead) {
            stream.Dispose();
            throw new InvalidOperationException("The OneNote payload stream factory returned a non-readable stream.");
        }
        return stream;
    }

    /// <summary>Materializes a defensive copy of the payload with a caller-provided size bound.</summary>
    public byte[] ToArray(long maxBytes) {
        if (maxBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxBytes));
        if (Length.HasValue && Length.Value > maxBytes) {
            throw new IOException("The OneNote payload exceeds the requested materialization limit.");
        }

        using (Stream stream = OpenRead())
        using (var output = new MemoryStream()) {
            var buffer = new byte[64 * 1024];
            long total = 0;
            while (true) {
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read <= 0) break;
                total += read;
                if (total > maxBytes) throw new IOException("The OneNote payload exceeds the requested materialization limit.");
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }
    }
}

/// <summary>
/// Base class for OneNote elements backed by binary payloads.
/// </summary>
public abstract class OneNoteBinaryElement : OneNoteElement {
    /// <summary>Original or synthesized file name.</summary>
    public string? FileName { get; set; }

    /// <summary>Media type when known.</summary>
    public string? MediaType { get; set; }

    /// <summary>Lazy payload handle.</summary>
    public OneNoteBinaryPayload? Payload { get; set; }

    internal OneNoteExtendedGuid? PayloadObjectId { get; set; }
    internal Guid? PayloadFileDataId { get; set; }
    internal string? PayloadFileExtension { get; set; }
}

/// <summary>
/// An image placed on a page or in an outline.
/// </summary>
public sealed class OneNoteImage : OneNoteBinaryElement {
    /// <inheritdoc />
    public override OneNoteElementKind Kind => OneNoteElementKind.Image;

    /// <summary>Alternative text.</summary>
    public string? AltText { get; set; }

    /// <summary>Source path recorded by OneNote.</summary>
    public string? SourcePath { get; set; }

    /// <summary>Optional hyperlink associated with the image.</summary>
    public string? Hyperlink { get; set; }

    /// <summary>
    /// Original image width in the half-inch increments used by the MS-ONE
    /// <c>PictureWidth</c> property.
    /// </summary>
    public double? WidthHalfInches { get; set; }

    /// <summary>
    /// Original image height in the half-inch increments used by the MS-ONE
    /// <c>PictureHeight</c> property.
    /// </summary>
    public double? HeightHalfInches { get; set; }

    internal OneNoteExtendedGuid? PictureContainerObjectId { get; set; }
    internal OneNoteExtendedGuid? WebPictureContainerObjectId { get; set; }
    internal bool PayloadUsesWebPictureContainer { get; set; }
}

/// <summary>
/// A file embedded in a OneNote page.
/// </summary>
public sealed class OneNoteEmbeddedFile : OneNoteBinaryElement {
    /// <inheritdoc />
    public override OneNoteElementKind Kind => OneNoteElementKind.EmbeddedFile;

    /// <summary>Original source path recorded by OneNote.</summary>
    public string? SourcePath { get; set; }
}

/// <summary>
/// Ink or handwriting content preserved from a OneNote page.
/// </summary>
public sealed class OneNoteInk : OneNoteBinaryElement {
    /// <inheritdoc />
    public override OneNoteElementKind Kind => OneNoteElementKind.Ink;

    /// <summary>Decoded strokes when the ink representation is understood.</summary>
    public IList<OneNoteInkStroke> Strokes { get; } = new List<OneNoteInkStroke>();
}

/// <summary>
/// A decoded ink stroke.
/// </summary>
public sealed class OneNoteInkStroke {
    /// <summary>Stroke points in source order.</summary>
    public IList<OneNoteInkPoint> Points { get; } = new List<OneNoteInkPoint>();

    /// <summary>Stroke color encoded as ARGB.</summary>
    public uint? ColorArgb { get; set; }

    /// <summary>Stroke width.</summary>
    public double? Width { get; set; }
}

/// <summary>
/// A point in a decoded ink stroke.
/// </summary>
public sealed class OneNoteInkPoint {
    /// <summary>Horizontal coordinate.</summary>
    public double X { get; set; }

    /// <summary>Vertical coordinate.</summary>
    public double Y { get; set; }

    /// <summary>Optional pen pressure.</summary>
    public double? Pressure { get; set; }
}

/// <summary>
/// Mathematical content and its best available projections.
/// </summary>
public sealed class OneNoteMath : OneNoteElement {
    /// <inheritdoc />
    public override OneNoteElementKind Kind => OneNoteElementKind.Math;

    /// <summary>Plain-text mathematical projection.</summary>
    public string? Text { get; set; }

    /// <summary>MathML projection when available.</summary>
    public string? MathMl { get; set; }

    /// <summary>LaTeX projection when available.</summary>
    public string? Latex { get; set; }

    /// <summary>Raw mathematical payload preserved for loss-aware writing.</summary>
    public OneNoteBinaryPayload? RawPayload { get; set; }
}

/// <summary>
/// Audio or video content referenced by a page.
/// </summary>
public sealed class OneNoteMedia : OneNoteBinaryElement {
    /// <inheritdoc />
    public override OneNoteElementKind Kind => OneNoteElementKind.Media;

    /// <summary>Recording duration when known.</summary>
    public TimeSpan? Duration { get; set; }

    /// <summary>Recording identity when known.</summary>
    public Guid? RecordingId { get; set; }

    /// <summary>Whether the recording contains audio or video.</summary>
    public OneNoteMediaKind RecordingKind { get; set; }

    /// <summary>Original source path recorded by OneNote.</summary>
    public string? SourcePath { get; set; }
}

/// <summary>OneNote recording media classification.</summary>
public enum OneNoteMediaKind {
    /// <summary>The recording kind is unavailable.</summary>
    Unknown = 0,
    /// <summary>An audio recording.</summary>
    Audio = 1,
    /// <summary>A video recording.</summary>
    Video = 2
}
