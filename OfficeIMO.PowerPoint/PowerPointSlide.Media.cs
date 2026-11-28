using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds an embedded video file to the slide.
        /// </summary>
        public PowerPointMedia AddVideo(string videoPath, string? posterImagePath = null, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null) {
            ArgumentNullException.ThrowIfNull(videoPath);
            if (!File.Exists(videoPath)) {
                throw new FileNotFoundException("Video file not found.", videoPath);
            }

            using FileStream videoStream = new(videoPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (posterImagePath == null) {
                return AddVideo(videoStream, GetVideoContentType(videoPath), Path.GetExtension(videoPath), left, top, width, height);
            }

            if (!File.Exists(posterImagePath)) {
                throw new FileNotFoundException("Poster image file not found.", posterImagePath);
            }

            using FileStream posterStream = new(posterImagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return AddVideo(videoStream, GetVideoContentType(videoPath), Path.GetExtension(videoPath), left, top, width, height,
                posterStream, GetImagePartType(posterImagePath));
        }

        /// <summary>
        ///     Adds an embedded video stream to the slide.
        /// </summary>
        public PowerPointMedia AddVideo(Stream video, string contentType, string extension, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null, Stream? posterImage = null,
            OfficeImageFormat posterImageType = OfficeImageFormat.Png) {
            return AddMedia(video, contentType, extension, PowerPointMediaKind.Video,
                (left ?? OfficeLength.Emus(0L)).ToEmus(), (top ?? OfficeLength.Emus(0L)).ToEmus(),
                (width ?? OfficeLength.Emus(3657600L)).ToEmus(), (height ?? OfficeLength.Emus(2057400L)).ToEmus(),
                posterImage, posterImageType);
        }

        /// <summary>
        ///     Adds an embedded audio file to the slide.
        /// </summary>
        public PowerPointMedia AddAudio(string audioPath, OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null,
            OfficeLength? height = null) {
            ArgumentNullException.ThrowIfNull(audioPath);
            if (!File.Exists(audioPath)) {
                throw new FileNotFoundException("Audio file not found.", audioPath);
            }

            using FileStream audioStream = new(audioPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return AddAudio(audioStream, GetAudioContentType(audioPath), Path.GetExtension(audioPath), left, top, width, height);
        }

        /// <summary>
        ///     Adds an embedded audio stream to the slide.
        /// </summary>
        public PowerPointMedia AddAudio(Stream audio, string contentType, string extension, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null) {
            return AddMedia(audio, contentType, extension, PowerPointMediaKind.Audio,
                (left ?? OfficeLength.Emus(0L)).ToEmus(), (top ?? OfficeLength.Emus(0L)).ToEmus(),
                (width ?? OfficeLength.Emus(914400L)).ToEmus(), (height ?? OfficeLength.Emus(914400L)).ToEmus());
        }

        /// <summary>Adds a linked audio frame without copying the payload into the package.</summary>
        public PowerPointMedia AddLinkedAudio(Uri source, OfficeLength? left = null,
            OfficeLength? top = null, OfficeLength? width = null, OfficeLength? height = null) =>
            AddLinkedMedia(source, PowerPointMediaKind.Audio,
                (left ?? OfficeLength.Emus(0L)).ToEmus(), (top ?? OfficeLength.Emus(0L)).ToEmus(),
                (width ?? OfficeLength.Emus(914400L)).ToEmus(), (height ?? OfficeLength.Emus(914400L)).ToEmus());

        /// <summary>Adds a linked video frame without copying the payload into the package.</summary>
        public PowerPointMedia AddLinkedVideo(Uri source, OfficeLength? left = null,
            OfficeLength? top = null, OfficeLength? width = null, OfficeLength? height = null) =>
            AddLinkedMedia(source, PowerPointMediaKind.Video,
                (left ?? OfficeLength.Emus(0L)).ToEmus(), (top ?? OfficeLength.Emus(0L)).ToEmus(),
                (width ?? OfficeLength.Emus(3657600L)).ToEmus(), (height ?? OfficeLength.Emus(2057400L)).ToEmus());

        private PowerPointMedia AddLinkedMedia(Uri source, PowerPointMediaKind kind,
            long left, long top, long width, long height) {
            ArgumentNullException.ThrowIfNull(source);
            if (!source.IsAbsoluteUri) throw new ArgumentException("A linked media URI must be absolute.", nameof(source));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);
            if (TryAddPackageMedia(null, null, null, source, kind, left, top, width, height, null, OfficeImageFormat.Png) is { } packageMedia) return packageMedia;
            uint shapeId = AllocateShapeId();
            SlidePart slidePart = SlidePart;
            string relationshipId = GetNextRelationshipId(slidePart);
            string relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" +
                (kind == PowerPointMediaKind.Audio ? "audio" : "video");
            var part = PartRef;
            uint timingId = GetNewMediaTimingId(part.Root);
            ExternalRelationship link = slidePart.AddExternalRelationship(relationshipType, source, relationshipId);
            ImagePart? poster = null;
            try {
                poster = AddGeneratedMediaPoster(kind);
                string name = GenerateUniqueName(kind == PowerPointMediaKind.Audio ? "Linked Audio" : "Linked Video");
                return AppendMediaFrame(part, kind, shapeId, name, relationshipId, relationshipId, slidePart.GetIdOfPart(poster),
                    true, left, top, width, height, timingId);
            } catch {
                if (poster != null) slidePart.DeletePart(poster);
                slidePart.DeleteReferenceRelationship(link);
                throw;
            }
        }

        private PowerPointMedia AddMedia(Stream media, string contentType, string extension, PowerPointMediaKind kind,
            long left, long top, long width, long height, Stream? posterImage = null,
            OfficeImageFormat posterImageType = OfficeImageFormat.Png) {
            ArgumentNullException.ThrowIfNull(media);
            if (!media.CanRead) {
                throw new ArgumentException("Media stream must be readable.", nameof(media));
            }
            ArgumentException.ThrowIfNullOrWhiteSpace(contentType);
            ArgumentException.ThrowIfNullOrWhiteSpace(extension);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);

            using MemoryStream? normalizedPoster = posterImage == null ? null
                : PowerPointImageNormalization.NormalizeHeicForEmbedding(posterImage, posterImageType);
            if (normalizedPoster != null)
                return AddMedia(media, contentType, extension, kind, left, top, width, height, normalizedPoster, OfficeImageFormat.Png);

            if (posterImage is { CanRead: false }) throw new ArgumentException("Image stream must be readable.", "image");
            if (TryAddPackageMedia(media, contentType, extension, null, kind, left, top, width, height, posterImage, posterImageType) is { } packageMedia) return packageMedia;
            uint shapeId = AllocateShapeId();
            SlidePart slidePart = SlidePart;
            var part = PartRef;
            uint timingId = GetNewMediaTimingId(part.Root);
            MediaDataPart mediaPart = CreateMediaDataPart(media, contentType, extension);
            DataPartReferenceRelationship? fileReference = null;
            MediaReferenceRelationship? playbackReference = null;
            ImagePart? posterPart = null;
            try {
                string fileId = GetNextRelationshipId(slidePart);
                fileReference = kind == PowerPointMediaKind.Audio
                    ? slidePart.AddAudioReferenceRelationship(mediaPart, fileId)
                    : slidePart.AddVideoReferenceRelationship(mediaPart, fileId);
                playbackReference = slidePart.AddMediaReferenceRelationship(mediaPart, GetNextRelationshipId(slidePart));
                posterPart = posterImage == null ? AddGeneratedMediaPoster(kind) : AddImagePartFromStream(posterImage, posterImageType);
                string name = GenerateUniqueName(kind == PowerPointMediaKind.Audio ? "Audio" : "Video");
                return AppendMediaFrame(part, kind, shapeId, name, fileId, playbackReference.Id, slidePart.GetIdOfPart(posterPart),
                    false, left, top, width, height, timingId);
            } catch {
                if (posterPart != null) slidePart.DeletePart(posterPart);
                if (playbackReference != null) slidePart.DeleteReferenceRelationship(playbackReference);
                if (fileReference != null) slidePart.DeleteReferenceRelationship(fileReference);
                ((PresentationDocument)slidePart.OpenXmlPackage).DeletePart(mediaPart);
                throw;
            }
        }

        private MediaDataPart CreateMediaDataPart(Stream media, string contentType, string extension) {
            if (SlidePart.OpenXmlPackage is not PresentationDocument document) {
                throw new InvalidOperationException("Slide is not attached to a presentation document.");
            }

            string normalizedExtension = extension.Trim().TrimStart('.');
            MediaDataPart mediaPart = document.CreateMediaDataPart(contentType, normalizedExtension);
            try {
                if (media.CanSeek) media.Position = 0;
                mediaPart.FeedData(media);
                return mediaPart;
            } catch {
                document.DeletePart(mediaPart);
                throw;
            }
        }

        private ImagePart AddGeneratedMediaPoster(PowerPointMediaKind kind) {
            byte[] poster = CreateGeneratedMediaPosterPng(kind);
            using MemoryStream stream = new(poster);
            return AddImagePartFromStream(stream, OfficeImageFormat.Png);
        }

        private static byte[] CreateGeneratedMediaPosterPng(PowerPointMediaKind kind) {
            OfficeRasterImage image = new(640, 360, OfficeColor.FromRgb(31, 41, 55));
            OfficeRasterCanvas canvas = new(image);
            canvas.FillRectangle(0D, 0D, 640D, 360D, OfficeColor.FromRgb(31, 41, 55));
            canvas.FillEllipse(248D, 92D, 144D, 144D, OfficeColor.FromRgba(249, 250, 251, 240));

            if (kind == PowerPointMediaKind.Audio) {
                canvas.FillRectangle(280D, 140D, 18D, 64D, OfficeColor.FromRgb(17, 24, 39));
                canvas.FillRectangle(311D, 118D, 18D, 86D, OfficeColor.FromRgb(17, 24, 39));
                canvas.FillRectangle(342D, 154D, 18D, 50D, OfficeColor.FromRgb(17, 24, 39));
            } else {
                canvas.FillPolygon(new[] {
                    new OfficePoint(296D, 130D),
                    new OfficePoint(296D, 198D),
                    new OfficePoint(358D, 164D)
                }, OfficeColor.FromRgb(17, 24, 39));
            }

            return OfficePngWriter.Encode(image);
        }

        private ImagePart AddImagePartFromStream(Stream image, OfficeImageFormat imageType) {
            ArgumentNullException.ThrowIfNull(image);
            if (!image.CanRead) {
                throw new ArgumentException("Image stream must be readable.", nameof(image));
            }

            using MemoryStream? normalized = PowerPointImageNormalization.NormalizeHeicForEmbedding(image, imageType);
            if (normalized != null) return AddImagePartFromStream(normalized, OfficeImageFormat.Png);

            PartTypeInfo partTypeInfo = imageType.ToPartTypeInfo();
            string imageExtension = PowerPointPartFactory.GetImageExtension(imageType);
            string imagePartUri = PowerPointPartFactory.GetIndexedPartUri(
                SlidePart.OpenXmlPackage,
                "ppt/media",
                "image",
                imageExtension,
                allowBaseWithoutIndex: false);
            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(
                SlidePart,
                partTypeInfo.ContentType,
                imagePartUri);
            try {
                if (image.CanSeek) image.Position = 0;
                imagePart.FeedData(image);
                return imagePart;
            } catch {
                SlidePart.DeletePart(imagePart);
                throw;
            }
        }

        private static string GetAudioContentType(string mediaPath) {
            return Path.GetExtension(mediaPath).ToLowerInvariant() switch {
                ".wav" => "audio/wav",
                ".wma" => "audio/x-ms-wma",
                ".ogg" or ".oga" => "audio/ogg",
                ".m4a" => "audio/mp4",
                ".mid" or ".midi" => "audio/midi",
                ".aiff" or ".aif" => "audio/aiff",
                _ => "audio/mpeg"
            };
        }

        private static string GetVideoContentType(string mediaPath) {
            return Path.GetExtension(mediaPath).ToLowerInvariant() switch {
                ".avi" => "video/x-msvideo",
                ".mpeg" => "video/mpeg",
                ".mpg" => "video/mpg",
                ".mov" => "video/quicktime",
                ".wmv" => "video/x-ms-wmv",
                ".ogv" => "video/ogg",
                _ => "video/mp4"
            };
        }
    }
}
