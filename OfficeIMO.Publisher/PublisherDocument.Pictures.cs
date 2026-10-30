namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Replaces one supported picture placement's embedded image with owned PNG or JPEG bytes. Other placements sharing its old asset remain unchanged.</summary>
    /// <remarks>Retains picture identity, page/master ownership, crop, bounds and drawing order. Original native assets remain available in Images. Fetch current page, picture and image snapshots after an edit. Failed validation, cancellation or projection leaves the document unchanged. Native PUB save-back is not provided.</remarks>
    public void ReplacePictureImage(uint pictureId, byte[] bytes, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        cancellationToken.ThrowIfCancellationRequested();
        var model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        Commit(model.ReplacePictureImage(pictureId, bytes, cancellationToken).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Sets a supported picture's bounds in page-local points. Negative offsets allow page-clipped bleed; width and height must be positive and finite.</summary>
    /// <remarks>Only pictures whose CanEdit property is true are supported. Retains image, crop, identity, ownership and drawing order, and recalculates text wrapping and page scenes. Old snapshots remain unchanged; failed or cancelled edits leave the document unchanged. Concurrent mutation is unsupported.</remarks>
    public void SetPictureBounds(uint pictureId, double x, double y, double width, double height, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!Finite(x)) throw new ArgumentOutOfRangeException(nameof(x));
        if (!Finite(y)) throw new ArgumentOutOfRangeException(nameof(y));
        if (!Finite(width) || width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (!Finite(height) || height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        if (!Finite(x + width)) throw new ArgumentOutOfRangeException(nameof(width));
        if (!Finite(y + height)) throw new ArgumentOutOfRangeException(nameof(height));
        var model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        Commit(model.SetPictureBounds(pictureId, x, y, width, height).Project(cancellationToken), cancellationToken);
    }
}
