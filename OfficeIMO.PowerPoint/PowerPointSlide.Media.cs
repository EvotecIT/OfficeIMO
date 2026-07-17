using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds an embedded video file to the slide.
        /// </summary>
        public PowerPointMedia AddVideo(string videoPath, string? posterImagePath = null, long left = 0L, long top = 0L,
            long width = 3657600L, long height = 2057400L) {
            if (videoPath == null) {
                throw new ArgumentNullException(nameof(videoPath));
            }
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
        public PowerPointMedia AddVideo(Stream video, string contentType, string extension, long left = 0L, long top = 0L,
            long width = 3657600L, long height = 2057400L, Stream? posterImage = null,
            ImagePartType posterImageType = ImagePartType.Png) {
            return AddMedia(video, contentType, extension, PowerPointMediaKind.Video, left, top, width, height, posterImage,
                posterImageType);
        }

        /// <summary>
        ///     Adds an embedded audio file to the slide.
        /// </summary>
        public PowerPointMedia AddAudio(string audioPath, long left = 0L, long top = 0L, long width = 914400L,
            long height = 914400L) {
            if (audioPath == null) {
                throw new ArgumentNullException(nameof(audioPath));
            }
            if (!File.Exists(audioPath)) {
                throw new FileNotFoundException("Audio file not found.", audioPath);
            }

            using FileStream audioStream = new(audioPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return AddAudio(audioStream, GetAudioContentType(audioPath), Path.GetExtension(audioPath), left, top, width, height);
        }

        /// <summary>
        ///     Adds an embedded audio stream to the slide.
        /// </summary>
        public PowerPointMedia AddAudio(Stream audio, string contentType, string extension, long left = 0L, long top = 0L,
            long width = 914400L, long height = 914400L) {
            return AddMedia(audio, contentType, extension, PowerPointMediaKind.Audio, left, top, width, height);
        }

        private PowerPointMedia AddMedia(Stream media, string contentType, string extension, PowerPointMediaKind kind,
            long left, long top, long width, long height, Stream? posterImage = null,
            ImagePartType posterImageType = ImagePartType.Png) {
            if (media == null) {
                throw new ArgumentNullException(nameof(media));
            }
            if (!media.CanRead) {
                throw new ArgumentException("Media stream must be readable.", nameof(media));
            }
            if (string.IsNullOrWhiteSpace(contentType)) {
                throw new ArgumentException("Media content type is required.", nameof(contentType));
            }
            if (string.IsNullOrWhiteSpace(extension)) {
                throw new ArgumentException("Media extension is required.", nameof(extension));
            }
            if (width <= 0) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (height <= 0) {
                throw new ArgumentOutOfRangeException(nameof(height));
            }

            MediaDataPart mediaPart = CreateMediaDataPart(media, contentType, extension);
            string fileReferenceId = GetNextRelationshipId(_slidePart);
            if (kind == PowerPointMediaKind.Audio) {
                _slidePart.AddAudioReferenceRelationship(mediaPart, fileReferenceId);
            } else {
                _slidePart.AddVideoReferenceRelationship(mediaPart, fileReferenceId);
            }
            string playbackReferenceId = GetNextRelationshipId(_slidePart);
            _slidePart.AddMediaReferenceRelationship(mediaPart, playbackReferenceId);

            ImagePart posterPart = posterImage == null
                ? AddGeneratedMediaPoster(kind)
                : AddImagePartFromStream(posterImage, posterImageType);
            string posterRelationshipId = _slidePart.GetIdOfPart(posterPart);

            string name = GenerateUniqueName(kind == PowerPointMediaKind.Audio ? "Audio" : "Video");
            uint shapeId = AllocateShapeId();
            Picture picture = CreateMediaPicture(kind, shapeId, name, fileReferenceId, playbackReferenceId,
                posterRelationshipId, left, top, width, height);

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(picture);
            EnsureMediaTiming(shapeId, kind);

            return TrackShape(new PowerPointMedia(picture, _slidePart, kind));
        }

        private MediaDataPart CreateMediaDataPart(Stream media, string contentType, string extension) {
            if (_slidePart.OpenXmlPackage is not PresentationDocument document) {
                throw new InvalidOperationException("Slide is not attached to a presentation document.");
            }

            string normalizedExtension = extension.Trim().TrimStart('.');
            MediaDataPart mediaPart = document.CreateMediaDataPart(contentType, normalizedExtension);
            if (media.CanSeek) {
                media.Position = 0;
            }
            mediaPart.FeedData(media);
            return mediaPart;
        }

        private ImagePart AddGeneratedMediaPoster(PowerPointMediaKind kind) {
            byte[] poster = CreateGeneratedMediaPosterPng(kind);
            using MemoryStream stream = new(poster);
            return AddImagePartFromStream(stream, ImagePartType.Png);
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

        private ImagePart AddImagePartFromStream(Stream image, ImagePartType imageType) {
            if (image == null) {
                throw new ArgumentNullException(nameof(image));
            }
            if (!image.CanRead) {
                throw new ArgumentException("Image stream must be readable.", nameof(image));
            }

            PartTypeInfo partTypeInfo = imageType.ToPartTypeInfo();
            string imageExtension = PowerPointPartFactory.GetImageExtension(imageType);
            string imagePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/media",
                "image",
                imageExtension,
                allowBaseWithoutIndex: false);
            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(
                _slidePart,
                partTypeInfo.ContentType,
                imagePartUri);
            if (image.CanSeek) {
                image.Position = 0;
            }
            imagePart.FeedData(image);
            return imagePart;
        }

        private static Picture CreateMediaPicture(PowerPointMediaKind kind, uint shapeId, string name,
            string fileReferenceId, string playbackReferenceId, string posterRelationshipId,
            long left, long top, long width, long height) {
            ApplicationNonVisualDrawingProperties appProperties = new();
            if (kind == PowerPointMediaKind.Audio) {
                appProperties.Append(new A.AudioFromFile { Link = fileReferenceId });
            } else {
                appProperties.Append(new A.VideoFromFile { Link = fileReferenceId });
            }

            P14.Media media = new() { Embed = playbackReferenceId };
            media.AddNamespaceDeclaration("p14", P14Namespace);
            ApplicationNonVisualDrawingPropertiesExtension extension =
                new() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };
            extension.Append(media);
            appProperties.Append(new ApplicationNonVisualDrawingPropertiesExtensionList(extension));

            NonVisualDrawingProperties drawingProperties = new() { Id = shapeId, Name = name };
            drawingProperties.Append(new A.HyperlinkOnClick { Id = string.Empty, Action = "ppaction://media" });

            return new Picture(
                new NonVisualPictureProperties(
                    drawingProperties,
                    new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    appProperties),
                new BlipFill(
                    new A.Blip { Embed = posterRelationshipId },
                    new A.Stretch(new A.FillRectangle())),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        }

        private void EnsureMediaTiming(uint shapeId, PowerPointMediaKind kind) {
            Timing timing = SlideRoot.Timing ??= new Timing();
            TimeNodeList timeNodeList = timing.GetFirstChild<TimeNodeList>() ?? timing.AppendChild(new TimeNodeList());
            ParallelTimeNode rootParallel = timeNodeList.GetFirstChild<ParallelTimeNode>() ?? timeNodeList.AppendChild(new ParallelTimeNode());
            CommonTimeNode rootTimeNode = rootParallel.GetFirstChild<CommonTimeNode>() ??
                rootParallel.AppendChild(new CommonTimeNode {
                    Id = 1U,
                    Duration = "indefinite",
                    Restart = TimeNodeRestartValues.Never,
                    NodeType = TimeNodeValues.TmingRoot
                });
            ChildTimeNodeList childNodes = rootTimeNode.GetFirstChild<ChildTimeNodeList>() ??
                rootTimeNode.AppendChild(new ChildTimeNodeList());

            OpenXmlCompositeElement mediaNode = kind == PowerPointMediaKind.Audio
                ? new Audio()
                : new Video();
            mediaNode.Append(
                new CommonMediaNode(
                    new CommonTimeNode(
                        new StartConditionList(new Condition { Delay = "indefinite" })) {
                        Id = GetNextTimingId(),
                        Fill = TimeNodeFillValues.Hold,
                        Display = false
                    },
                    new TargetElement(new ShapeTarget { ShapeId = shapeId.ToString(System.Globalization.CultureInfo.InvariantCulture) })) {
                    Volume = 80000
                });
            childNodes.Append(mediaNode);
        }

        private uint GetNextTimingId() {
            uint maxId = 0;
            foreach (CommonTimeNode node in SlideRoot.Descendants<CommonTimeNode>()) {
                uint? id = node.Id?.Value;
                if (id.HasValue && id.Value > maxId) {
                    maxId = id.Value;
                }
            }

            return maxId + 1;
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
