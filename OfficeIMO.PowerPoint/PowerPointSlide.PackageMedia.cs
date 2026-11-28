using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private PowerPointMedia? TryAddPackageMedia(Stream? media, string? contentType, string? extension, Uri? source,
            PowerPointMediaKind kind, long left, long top, long width, long height, Stream? poster, OfficeImageFormat posterFormat) {
            string? suffix = extension?.Trim().TrimStart('.');
            if (source == null && (suffix is not { Length: > 0 } || suffix.Any(character => !IsMediaSuffixCharacter(character))
                || !HasSimpleMediaContentType(contentType)
                || media is not { CanSeek: true } || media.Length > PowerPointMedia.MaximumOwnedPayloadBytes)) return null;
            if (OwnerDocument is not { } document || !TryGetEditablePackageShapeTree(out _, out _) || !TryGetPackageMaxShapeId(out _)) return null;
            string baseName = kind == PowerPointMediaKind.Audio ? "Audio" : "Video";
            string name = GenerateUniqueName(source == null ? baseName : "Linked " + baseName);
            PowerPointPartRef part = PartRef;
            if (part.IsSdk) return null;
            uint timingId = GetNewMediaTimingId(part.Root);

            byte[]? mediaBytes = null;
            if (media != null) {
                if (media.CanSeek) media.Position = 0;
                mediaBytes = OfficeStreamReader.ReadRemainingBytes(media);
            }
            byte[] posterBytes;
            if (poster == null) {
                posterBytes = CreateGeneratedMediaPosterPng(kind);
            } else {
                if (poster.CanSeek) poster.Position = 0;
                posterBytes = OfficeStreamReader.ReadRemainingBytes(poster);
            }
            string posterContentType = PowerPointImageContentTypes.For(posterFormat);
            string posterExtension = PowerPointPartFactory.GetImageExtension(posterFormat);
            uint shapeId = AllocateShapeId();
            string referenceType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + (kind == PowerPointMediaKind.Audio ? "audio" : "video");
            var (fileId, playbackId, posterId) = document.AddPackageMediaParts(part, referenceType, source, mediaBytes, contentType,
                suffix, posterBytes, posterContentType, posterExtension);
            return AppendMediaFrame(part, kind, shapeId, name, fileId, playbackId, posterId, source != null, left, top, width, height, timingId);
        }

        private static bool IsMediaSuffixCharacter(char character) => character is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9';

        // The SDK validates custom and parameterized MIME types. The owned path handles ordinary OPC types.
        private static bool HasSimpleMediaContentType(string? contentType) {
            if (contentType == null) return false;
            int slash = contentType.IndexOf('/');
            if (slash <= 0 || slash == contentType.Length - 1) return false;
            for (int index = 0; index < contentType.Length; index++) {
                char character = contentType[index];
                if (index != slash && !IsMediaSuffixCharacter(character) && character is not ('+' or '-' or '.')) return false;
            }
            return true;
        }

        private PowerPointMedia AppendMediaFrame(PowerPointPartRef part, PowerPointMediaKind kind, uint shapeId,
            string name, string fileId, string playbackId, string posterId, bool linked,
            long left, long top, long width, long height, uint timingId) {
            const int p = OfficeXmlNamespaces.PresentationML;
            const int a = OfficeXmlNamespaces.DrawingML;
            const int r = OfficeXmlNamespaces.Relationships;
            OfficeXmlNode tree = EnsureCommonSlideData(part).GetOrAddChild(p, "spTree", PCommonSlideDataView.Order);
            OfficeXmlNode picture = tree.CreateElement(p, "pic");
            OfficeXmlNode nonVisual = AppendMediaElement(picture, p, "nvPicPr");
            OfficeXmlNode drawing = AppendMediaElement(nonVisual, p, "cNvPr");
            drawing.SetUInt32Attribute("id", shapeId);
            drawing.SetAttribute("name", name);
            OfficeXmlNode click = AppendMediaElement(drawing, a, "hlinkClick");
            click.SetAttribute(r, "id", string.Empty);
            click.SetAttribute("action", "ppaction://media");
            OfficeXmlNode pictureProperties = AppendMediaElement(nonVisual, p, "cNvPicPr");
            AppendMediaElement(pictureProperties, a, "picLocks").SetAttribute("noChangeAspect", "1");
            OfficeXmlNode application = AppendMediaElement(nonVisual, p, "nvPr");
            string reference = kind == PowerPointMediaKind.Audio ? "audioFile" : "videoFile";
            AppendMediaElement(application, a, reference).SetAttribute(r, "link", fileId);
            OfficeXmlNode extensions = AppendMediaElement(application, p, "extLst");
            OfficeXmlNode extension = AppendMediaElement(extensions, p, "ext");
            extension.SetAttribute("uri", PowerPointMedia.MediaExtensionUri);
            OfficeXmlNode media = extension.Append(extension.CreateElement(P14Namespace, "media", "p14"));
            media.DeclareNamespace("p14", P14Namespace);
            media.SetAttribute(r, linked ? "link" : "embed", playbackId);
            OfficeXmlNode fill = AppendMediaElement(picture, p, "blipFill");
            AppendMediaElement(fill, a, "blip").SetAttribute(r, "embed", posterId);
            AppendMediaElement(AppendMediaElement(fill, a, "stretch"), a, "fillRect");
            OfficeXmlNode properties = AppendMediaElement(picture, p, "spPr");
            OfficeXmlNode transform = AppendMediaElement(properties, a, "xfrm");
            OfficeXmlNode offset = AppendMediaElement(transform, a, "off");
            offset.SetInt64Attribute("x", left);
            offset.SetInt64Attribute("y", top);
            OfficeXmlNode size = AppendMediaElement(transform, a, "ext");
            size.SetInt64Attribute("cx", width);
            size.SetInt64Attribute("cy", height);
            OfficeXmlNode geometry = AppendMediaElement(properties, a, "prstGeom");
            geometry.SetAttribute("prst", "rect");
            AppendMediaElement(geometry, a, "avLst");
            AppendMediaTiming(part.Root, shapeId, kind, timingId);
            picture = tree.Append(picture);
            part.MarkEdited();
            return TrackShape((PowerPointMedia)WrapSlideShape(picture, part)!);
        }

        private static OfficeXmlNode AppendMediaElement(OfficeXmlNode parent, int ns, string localName) =>
            parent.Append(parent.CreateElement(ns, localName));

        private static uint GetNewMediaTimingId(OfficeXmlNode slide) {
            uint nextId = GetNextTimingId(slide);
            OfficeXmlNode root = slide.Child(TimingNamespace, "timing").Child(TimingNamespace, "tnLst")
                .Child(TimingNamespace, "par").Child(TimingNamespace, "cTn");
            return nextId == 1 && root.IsNone ? 2 : nextId;
        }

        private static void AppendMediaTiming(OfficeXmlNode slide, uint shapeId, PowerPointMediaKind kind, uint id) {
            OfficeXmlNode timing = slide.Child(TimingNamespace, "timing");
            if (timing.IsNone) timing = InsertSlideTiming(slide);
            OfficeXmlNode list = timing.Child(TimingNamespace, "tnLst");
            if (list.IsNone) list = AppendTimingChild(timing, "tnLst");
            OfficeXmlNode parallel = list.Child(TimingNamespace, "par");
            if (parallel.IsNone) parallel = AppendTimingChild(list, "par");
            OfficeXmlNode root = parallel.Child(TimingNamespace, "cTn");
            if (root.IsNone) {
                root = AppendTimingChild(parallel, "cTn");
                root.SetUInt32Attribute("id", 1U);
                root.SetAttribute("dur", "indefinite");
                root.SetAttribute("restart", "never");
                root.SetAttribute("nodeType", "tmRoot");
            }
            OfficeXmlNode children = root.Child(TimingNamespace, "childTnLst");
            if (children.IsNone) children = AppendTimingChild(root, "childTnLst");
            OfficeXmlNode media = AppendTimingChild(children, kind == PowerPointMediaKind.Audio ? "audio" : "video");
            OfficeXmlNode common = AppendTimingChild(media, "cMediaNode");
            common.SetInt32Attribute("vol", 80000);
            OfficeXmlNode action = AppendTimingChild(common, "cTn");
            action.SetUInt32Attribute("id", id);
            action.SetAttribute("fill", "hold");
            action.SetAttribute("display", "0");
            AppendTimingChild(AppendTimingChild(action, "stCondLst"), "cond").SetAttribute("delay", "indefinite");
            OfficeXmlNode target = AppendTimingChild(AppendTimingChild(common, "tgtEl"), "spTgt");
            target.SetUInt32Attribute("spid", shapeId);
        }
    }
}
