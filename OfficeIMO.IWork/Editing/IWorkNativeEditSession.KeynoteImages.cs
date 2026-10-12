using System.Threading;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

public sealed partial class IWorkNativeEditSession {
    /// <summary>Moves a direct, unlocked, unrotated and unmasked Keynote image within its slide canvas.</summary>
    /// <remarks>Only its existing position floats change. IDs, references, dimensions, payload length and all other bytes remain intact.
    /// The identity must resolve to an image on this source's slides. Nested Index.zip editing is unsupported.</remarks>
    public void MoveKeynoteImage(IWorkObjectIdentity identity, float leftPoints, float topPoints,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (identity == null) throw new ArgumentNullException(nameof(identity));
        IWorkCreationValidation.ValidateFinite(leftPoints, nameof(leftPoints));
        IWorkCreationValidation.ValidateFinite(topPoints, nameof(topPoints));
        if (leftPoints < 0) throw new ArgumentOutOfRangeException(nameof(leftPoints));
        if (topPoints < 0) throw new ArgumentOutOfRangeException(nameof(topPoints));
        UpdateKeynoteImage(identity, leftPoints, topPoints, null, null, cancellationToken);
    }

    /// <summary>Resizes a direct, unlocked, unrotated and unmasked Keynote image within its slide canvas.</summary>
    /// <remarks>Dimensions are positive finite point sizes and may change the aspect ratio. Only the existing size floats change;
    /// position, intrinsic image pixels, IDs, references and all other bytes remain intact. Moves and resizes compose within
    /// this session. Nested Index.zip editing is unsupported, and edited saves require explicit visual-cache acceptance.</remarks>
    public void ResizeKeynoteImage(IWorkObjectIdentity identity, float widthPoints, float heightPoints,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (identity == null) throw new ArgumentNullException(nameof(identity));
        IWorkCreationValidation.ValidatePositive(widthPoints, nameof(widthPoints));
        IWorkCreationValidation.ValidatePositive(heightPoints, nameof(heightPoints));
        UpdateKeynoteImage(identity, null, null, widthPoints, heightPoints, cancellationToken);
    }

    private void UpdateKeynoteImage(IWorkObjectIdentity identity, float? leftPoints, float? topPoints,
        float? widthPoints, float? heightPoints, CancellationToken cancellationToken) {
        if (_source.Kind != IWorkDocumentKind.Keynote) throw new InvalidOperationException("Image editing requires a Keynote source.");
        if (_source.ContainerEntries.Any(entry => entry.Path == "Index.zip"))
            throw new NotSupportedException("Edits to nested Index.zip sources are not supported.");
        IWorkSourceDocument view = _source.WithCancellation(cancellationToken);
        IWorkKeynoteProjection projection = view.ReadKeynote();
        var matches = projection.Slides.SelectMany(slide => slide.Images.Select(image => (Slide: slide, Image: image)))
            .Where(item => SameIdentity(item.Image.SourceIdentity, identity)).Take(2).ToArray();
        if (matches.Length != 1) throw new ArgumentException("The identity must resolve to exactly one direct image in this source.", nameof(identity));
        var selected = matches[0]; var image = selected.Image;
        if (image.HasMask || image.Geometry == null || image.Geometry.RotationDegrees != 0 || projection.SlideSize == null)
            throw new NotSupportedException("Image editing requires complete unrotated, unmasked geometry and a known slide canvas.");
        IWorkArchiveRecord? record = view.Index.Find(identity.RecordIdentifier);
        if (record == null || record.MessageType != 3005 || record.PayloadIndex != 0 || !SameIdentity(new(record), identity))
            throw new ArgumentException("The identity does not identify a primary image record in this source.", nameof(identity));
        var message = view.Index.Message(record);
        var drawable = IWorkObjectIndex.TryGetMessage(message, 1, out bool malformed);
        var parent = drawable == null ? null : view.Index.Dereference(drawable, 2);
        if (malformed || drawable == null || parent?.Identifier != selected.Slide.SourceIdentity?.RecordIdentifier
            || drawable.FieldCount(5) > 1 || drawable.HasUnexpectedWireKind(5, IWorkWireKind.Varint)
            || drawable.GetUnsigned(5) is > 0)
            throw new NotSupportedException("Only unlocked direct slide images are supported.");
        byte[] current = _payloads.TryGetValue(record, out byte[]? pending) ? pending : record.Payload;
        var currentMessage = ReferenceEquals(current, record.Payload) ? message : IWorkProtobuf.Parse(current, view.Options);
        var currentDrawable = IWorkObjectIndex.TryGetMessage(currentMessage, 1, out bool currentMalformed);
        bool geometryComplete = false;
        IWorkGeometry? geometry = currentDrawable == null ? null
            : IWorkDrawingReader.ReadGeometry(currentDrawable, out geometryComplete, requirePositiveSize: true);
        if (currentMalformed || !geometryComplete || geometry == null)
            throw new NotSupportedException("The pending image geometry must remain complete.");
        double left = leftPoints.HasValue ? leftPoints.Value : geometry.LeftPoints;
        double top = topPoints.HasValue ? topPoints.Value : geometry.TopPoints;
        double width = widthPoints.HasValue ? widthPoints.Value : geometry.WidthPoints;
        double height = heightPoints.HasValue ? heightPoints.Value : geometry.HeightPoints;
        if (left < 0 || top < 0)
            throw new NotSupportedException("The image frame must have nonnegative coordinates.");
        if (left + width > projection.SlideSize.WidthPoints)
            throw new ArgumentOutOfRangeException(widthPoints.HasValue ? nameof(widthPoints) : nameof(leftPoints),
                "The image must fit within the slide canvas.");
        if (top + height > projection.SlideSize.HeightPoints)
            throw new ArgumentOutOfRangeException(heightPoints.HasValue ? nameof(heightPoints) : nameof(topPoints),
                "The image must fit within the slide canvas.");
        byte[] updated = (byte[])current.Clone();
        if (leftPoints.HasValue) {
            IWorkPreservingWireEdit.ReplaceFloat(updated, new[] { 1, 1, 1, 1 }, leftPoints.Value, view.Options, cancellationToken);
            IWorkPreservingWireEdit.ReplaceFloat(updated, new[] { 1, 1, 1, 2 }, topPoints!.Value, view.Options, cancellationToken);
        }
        if (widthPoints.HasValue) {
            IWorkPreservingWireEdit.ReplaceFloat(updated, new[] { 1, 1, 2, 1 }, widthPoints.Value, view.Options, cancellationToken);
            IWorkPreservingWireEdit.ReplaceFloat(updated, new[] { 1, 1, 2, 2 }, heightPoints!.Value, view.Options, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (updated.SequenceEqual(record.Payload)) _payloads.Remove(record); else _payloads[record] = updated;
    }

    private static bool SameIdentity(IWorkObjectIdentity? left, IWorkObjectIdentity right) => left != null
        && left.RecordIdentifier == right.RecordIdentifier && left.MessageType == right.MessageType
        && left.EntryPath == right.EntryPath && left.PayloadIndex == right.PayloadIndex;
}
