using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherModelState {
    internal PublisherModelState SetPictureCrop(uint pictureId, OfficeImageSourceCrop crop) =>
        WithPublication(_publication.SetCrop(_publication.EditablePicture(pictureId), crop));

    internal PublisherModelState SetPictureBounds(uint pictureId, double x, double y, double width, double height) {
        PublisherPicture picture = _publication.EditablePicture(pictureId);
        return WithPublication(_publication.SetBounds(picture, x, y, width, height));
    }

    internal PublisherModelState ReplacePictureImage(uint pictureId, byte[] bytes, CancellationToken token) {
        PublisherPicture picture = _publication.EditablePicture(pictureId);
        var payload = ValidatePictureBytes(bytes, token);
        PublisherPublicationState candidate = _publication.ReplaceImage(picture, payload.Bytes, payload.ContentType);
        if (candidate.AdditionalImageBytes > _decoded.Options.MaximumTotalImageBytes - _decoded.ImageBytes)
            throw new ArgumentException("Replacement image exceeds the configured publication image byte limit.", nameof(bytes));
        return WithPublication(candidate);
    }

    internal PublisherModelState AddPicture(uint pageId, byte[] bytes, double x, double y, double width, double height,
        CancellationToken token, out uint pictureId) {
        token.ThrowIfCancellationRequested();
        PublisherPageState page = _publication.PicturePage(pageId, x, y, width, height);
        if ((long)SourceObjectCount + _publication.AuthoredObjectCount >= _decoded.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher current object limit exceeded.");
        var payload = ValidatePictureBytes(bytes, token);
        if (payload.Decode.FrameCount != 1 || payload.Decode.IsAnimated || payload.Decode.SourceOrientation != OfficeImageOrientation.Normal)
            throw new ArgumentException("New pictures require a single static PNG or JPEG with normal encoded orientation.", nameof(bytes));
        PublisherPublicationState candidate = _publication.AddPicture(page, payload.Bytes, payload.ContentType, x, y, width, height, out pictureId);
        if (candidate.AdditionalImageBytes > _decoded.Options.MaximumTotalImageBytes - _decoded.ImageBytes)
            throw new ArgumentException("New picture exceeds the configured publication image byte limit.", nameof(bytes));
        return WithPublication(candidate);
    }

    // Addition and replacement share content validation and byte ownership.
    // Creation's static/orientation policy does not narrow existing replacements.
    private (byte[] Bytes, string ContentType, OfficeRasterDecodeInfo Decode) ValidatePictureBytes(byte[] bytes, CancellationToken token) {
        if (bytes.Length == 0 || bytes.Length > _decoded.Options.MaximumImageBytes)
            throw new ArgumentException("Picture bytes must fit the configured per-image limit.", nameof(bytes));
        byte[] owned = (byte[])bytes.Clone();
        if (!OfficeImageReader.TryIdentifyByContent(owned, null, token, out OfficeImageInfo info)
            || info.Format is not (OfficeImageFormat.Png or OfficeImageFormat.Jpeg))
            throw new ArgumentException("Pictures require PNG or JPEG content.", nameof(bytes));
        if (info.Width <= 0 || info.Height <= 0 || (long)info.Width * info.Height > _decoded.Options.MaximumRasterPixels)
            throw new ArgumentException("Picture dimensions exceed the configured pixel limit.", nameof(bytes));
        var options = new OfficeRasterDecodeOptions {
            MaximumEncodedBytes = Math.Min(_decoded.Options.MaximumImageBytes, 128 * 1024 * 1024),
            MaximumDecodedPixels = Math.Min(_decoded.Options.MaximumRasterPixels, 50_000_000),
            MaximumInspectionWorkPixels = Math.Min(_decoded.Options.MaximumImageProcessingPixels, 50_000_000),
            CancellationToken = token
        };
        if (!OfficeRasterImageDecoder.TryDecode(owned, options, out OfficeRasterImage? raster, out OfficeRasterDecodeInfo decoded) || raster == null)
            throw new ArgumentException("Picture content is invalid or exceeds the configured decoding bounds.", nameof(bytes));
        token.ThrowIfCancellationRequested();
        return (owned, info.MimeType, decoded);
    }
}
