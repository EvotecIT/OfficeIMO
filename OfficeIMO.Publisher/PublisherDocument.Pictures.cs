using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Sets normalized source-edge crop fractions on one supported picture without changing its frame or embedded image.</summary>
    /// <remarks>The crop must leave positive visible width and height. Retains identity, ownership, image bytes, bounds, paint and drawing order; other placements and old snapshots remain unchanged. Failed validation, cancellation or projection is atomic. Native source saving rounds to signed 16.16 fractions and reports approximation; a crop whose rounded area collapses refuses native output.</remarks>
    public void SetPictureCrop(uint pictureId, OfficeImageSourceCrop crop, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!crop.HasVisibleSourceArea) throw new ArgumentOutOfRangeException(nameof(crop), "The crop must leave a visible source area.");
        var model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.SetPictureCrop(pictureId, crop).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Appends an owned static PNG or JPEG picture to a qualified document or master page and returns its committed immutable snapshot. Bounds use physical page-local lengths and positive dimensions.</summary>
    /// <remarks>The detected content must have one static frame and normal encoded orientation. Bytes are copied; each call creates a separate embedded asset. The picture is uncropped, has no fill or outline, effects, transform, story or wrap references, and is appended in drawing order. Negative offsets allow clipped bleed. Existing picture edits remain available. Validation, limits, cancellation or projection failure consumes neither the object nor image identifier and leaves the publication unchanged. This is in-memory authoring, without native PUB saving.</remarks>
    public PublisherPicture AddImage(uint pageId, byte[] bytes, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(bytes);
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width.ToEmus(), nameof(width));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height.ToEmus(), nameof(height));
        var model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        PublisherDocument candidate = model.AddPicture(pageId, bytes, x.ToPoints(), y.ToPoints(), width.ToPoints(), height.ToPoints(), cancellationToken, out uint id).Project(cancellationToken);
        PublisherPicture result = candidate.Pages.Concat(candidate.MasterPages).SelectMany(page => page.Pictures).Single(picture => picture.Id == id);
        Commit(candidate, cancellationToken);
        return result;
    }

    /// <summary>Appends an uncropped picture that shares an existing document image asset. Bounds use physical page-local lengths.</summary>
    /// <remarks>The image identifier comes from this document's Images collection. Only single static PNG or JPEG images with normal encoded orientation are admitted. No asset is copied or reserved. The new picture has independent bounds, crop and paint; replacing its image leaves other placements unchanged. Failed validation, cancellation or projection leaves the publication and identifiers unchanged. This is in-memory authoring, without native PUB saving.</remarks>
    public PublisherPicture AddImage(uint pageId, int imageId, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width.ToEmus(), nameof(width));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height.ToEmus(), nameof(height));
        var model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        PublisherDocument candidate = model.AddPicture(pageId, imageId, x.ToPoints(), y.ToPoints(), width.ToPoints(), height.ToPoints(), cancellationToken, out uint id).Project(cancellationToken);
        PublisherPicture result = candidate.Pages.Concat(candidate.MasterPages).SelectMany(page => page.Pictures).Single(picture => picture.Id == id);
        Commit(candidate, cancellationToken);
        return result;
    }

    /// <summary>Replaces one supported picture placement's embedded image with owned PNG or JPEG bytes. Other placements sharing its old asset remain unchanged.</summary>
    /// <remarks>Retains picture identity, page/master ownership, crop, bounds and drawing order. Original native assets remain available in Images. Fetch current page, picture and image snapshots after an edit. Failed validation, cancellation or projection leaves the document unchanged. Qualified imported pictures support bounded native save-back with explicit loss-policy allowance; unsafe source bindings refuse output. Asset identifiers are document-scoped and may be remapped on native reload.</remarks>
    public void ReplacePictureImage(uint pictureId, byte[] bytes, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(bytes);
        cancellationToken.ThrowIfCancellationRequested();
        var model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.ReplacePictureImage(pictureId, bytes, cancellationToken).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Sets a supported picture's physical page-local bounds. Negative offsets allow page-clipped bleed; width and height must be positive.</summary>
    /// <remarks>Only pictures whose CanEdit property is true are supported. Retains image, crop, identity, ownership and drawing order, and recalculates text wrapping and page scenes. Old snapshots remain unchanged; failed or cancelled edits leave the document unchanged. Concurrent mutation is unsupported.</remarks>
    public void SetPictureBounds(uint pictureId, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width.ToEmus(), nameof(width));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height.ToEmus(), nameof(height));
        var model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.SetPictureBounds(pictureId, x.ToPoints(), y.ToPoints(), width.ToPoints(), height.ToPoints(), cancellationToken).Project(cancellationToken), cancellationToken);
    }
}
