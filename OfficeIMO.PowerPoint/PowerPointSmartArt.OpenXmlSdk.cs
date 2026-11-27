using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The SDK side of a SmartArt diagram: its constructor over an SDK graphic frame, the bytes of its diagram parts
    /// while the SDK holds them, the canonical layout definition the slide builds with SDK types, and the SDK group the
    /// slide's shape readers take for the cached drawing. Everything else reads and edits the diagram parts through
    /// <see cref="PowerPointPartRef"/>; this file moves into the bridge with the rest of the SDK side.
    /// </summary>
    public partial class PowerPointSmartArt {
        internal PowerPointSmartArt(GraphicFrame graphicFrame, SlidePart slidePart) : base(graphicFrame) {
            _initialSlidePart = PowerPointPartRef.FromSdk(slidePart);
        }

        private static byte[] ReadSdkPartBytes(object part) {
            try {
                using Stream stream = ((OpenXmlPart)part).GetStream(FileMode.Open, FileAccess.Read);
                using var buffer = new MemoryStream();
                stream.CopyTo(buffer);
                return buffer.ToArray();
            } catch (OpenXmlPackageException exception) {
                throw new InvalidDataException(exception.Message, exception);
            }
        }

        /// <summary>Parses an SDK part's XML straight from its stream.</summary>
        private static System.Xml.Linq.XDocument LoadSdkPartXml(object part, System.Xml.Linq.LoadOptions options) {
            Stream stream;
            try {
                stream = ((OpenXmlPart)part).GetStream(FileMode.Open, FileAccess.Read);
            } catch (OpenXmlPackageException exception) {
                throw new InvalidDataException(exception.Message, exception);
            }
            using (stream) return PowerPointXmlReader.LoadPackagePartXml(stream, options);
        }

        /// <summary>Saves the document straight into an SDK part's stream.</summary>
        private static void SaveSdkPartXml(object part, System.Xml.Linq.XDocument document) {
            using Stream stream = ((OpenXmlPart)part).GetStream(FileMode.Create, FileAccess.Write);
            document.Save(stream);
        }

        private static void WriteSdkPartBytes(object part, byte[] bytes) {
            using Stream stream = ((OpenXmlPart)part).GetStream(FileMode.Create, FileAccess.Write);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>The part as the SDK holds it, opening the SDK when it is not open yet.</summary>
        private static object? OpenSdkPart(PowerPointPartRef part) {
            if (part.Current.SdkPart is { } open) return open;
            if (part.Document is { } document) _ = document.OpenXmlDocument;
            return part.Current.SdkPart;
        }

        /// <summary>The layout definition the slide writes for a new diagram of this family, as markup.</summary>
        private static string CanonicalLayoutDefinitionXml(PowerPointSmartArtType type, int nodeCount, double aspectRatio) =>
            PowerPointSlide.CreateSmartArtLayoutDefinition(type, nodeCount, aspectRatio).OuterXml;

        /// <summary>
        /// The cached drawing as a group the slide's shape readers take: its markup read into an SDK group whose
        /// relationships resolve against the drawing part, which opens the SDK when it is not open yet.
        /// </summary>
        private PowerPointGroupShape CreateCachedDrawingGroup(PowerPointSlide slide, byte[] groupMarkup, PowerPointPartRef drawingPart) {
            var part = (OpenXmlPart)(OpenSdkPart(drawingPart)
                ?? throw new InvalidOperationException("The SmartArt cached drawing part was not found."));
            var group = new GroupShape(System.Text.Encoding.UTF8.GetString(groupMarkup));
            return (PowerPointGroupShape)new PowerPointGroupShape(group, part).AttachTo(slide, part);
        }
    }
}
