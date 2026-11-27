using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    /// <summary>
    /// What the writer's core still needs the Open XML SDK for, kept in one place so the rest reads the captured
    /// package view: wrapping layout, master and group shapes (the engine wraps those only over SDK elements), the
    /// callees that still take SDK elements, and the SDK-typed entry points that callers not yet passing a view use.
    /// Each forwarder names its caller; the file goes away once those callers and callees read the view.
    /// </summary>
    internal static partial class LegacyPptWriter {
        // ---- Shape wrappers ---------------------------------------------------------------------------------------

        /// <summary>Opens the SDK, which the engine needs to wrap layout, master and group shapes.</summary>
        private static void EnsureShapeWrappersCanBind(PowerPointDocument presentation) =>
            _ = presentation.OpenXmlDocument;

        /// <summary>
        /// Wraps a shape-tree element of the live part <paramref name="ownerPart"/> (an element of the view or of the
        /// live document), or null for content the writer does not encode.
        /// </summary>
        private static PowerPointShape? WrapShapeForWrite(PowerPointDocument presentation,
            PowerPointPartRef ownerPart, OfficeXmlNode element) {
            OpenXmlElement sdkElement = LiveSdkElement<OpenXmlElement>(presentation, ownerPart.Name, element)
                ?? throw new InvalidOperationException("A shape-tree element is missing from the live document.");
            OpenXmlPartContainer part = ownerPart.Current.SdkPart as OpenXmlPartContainer
                ?? throw new InvalidOperationException("The part that holds a shape is missing from the live document.");
            return WrapShapeForWrite(sdkElement, part);
        }

        private static PowerPointShape? WrapShapeForWrite(
            OpenXmlElement element, OpenXmlPartContainer ownerPart) =>
            element switch {
                P.Shape shape when shape.TextBody != null =>
                    new PowerPointTextBox(shape, ownerPart),
                P.Shape shape => new PowerPointAutoShape(shape),
                P.ConnectionShape connection =>
                    new PowerPointConnectionShape(connection),
                P.Picture picture when ownerPart is SlidePart slidePart
                    && PowerPointMedia.TryGetMediaKind(picture,
                        out PowerPointMediaKind kind) =>
                    new PowerPointMedia(picture, slidePart, kind),
                P.Picture picture when PowerPointMedia.TryGetMediaKind(
                    picture, out _) => null,
                P.Picture picture => new PowerPointImage(picture,
                    ownerPart),
                P.GroupShape nested => new PowerPointGroupShape(nested,
                    ownerPart),
                P.GraphicFrame frame when frame.Graphic?.GraphicData?
                    .GetFirstChild<A.Table>() != null =>
                    new PowerPointTable(frame, ownerPart),
                P.GraphicFrame frame when frame.Graphic?.GraphicData?
                    .GetFirstChild<C.ChartReference>() != null =>
                    new PowerPointChart(frame, ownerPart),
                P.GraphicFrame frame when frame.Graphic?.GraphicData?
                    .GetFirstChild<Dgm.RelationshipIds>() != null
                    && ownerPart is SlidePart slidePart =>
                    new PowerPointSmartArt(frame, slidePart),
                P.GraphicFrame frame when frame.Graphic?.GraphicData?
                    .GetFirstChild<P.OleObject>() != null
                    && ownerPart is SlidePart slidePart =>
                    new PowerPointOleObject(frame, slidePart),
                _ => null
            };

        internal static IReadOnlyList<PowerPointShape>
            ReadGroupChildrenForWrite(PowerPointGroupShape group,
                out string? unsupportedReason) {
            ArgumentNullException.ThrowIfNull(group);
            unsupportedReason = null;
            if (group.OwnerPart == null) {
                unsupportedReason = "The group shape has no owning package part for its nested content.";
                return Array.Empty<PowerPointShape>();
            }
            var result = new List<PowerPointShape>();
            foreach (OpenXmlElement element in group.GroupShape.ChildElements) {
                if (element is P.NonVisualGroupShapeProperties
                    or P.GroupShapeProperties) continue;
                PowerPointShape? child = WrapShapeForWrite(element,
                    group.OwnerPart);
                if (child == null) {
                    unsupportedReason = $"The group contains '{element.LocalName}' content that is not yet encoded by the native binary writer.";
                    return result;
                }
                if (group.OwnerSlide != null) child.AttachTo(group.OwnerSlide, group.OwnerPart);
                result.Add(child);
            }
            if (result.Count == 0) {
                unsupportedReason = "Binary PowerPoint groups must contain at least one drawable child.";
            }
            return result;
        }

        // ---- Live SDK elements for callees that still take them -----------------------------------------------

        /// <summary>
        /// The live document's SDK element at the place of <paramref name="node"/> (of the view part
        /// <paramref name="partName"/>, or already of the live document), opening the SDK when it is not open.
        /// </summary>
        private static T? LiveSdkElement<T>(PowerPointDocument? presentation, string partName,
            OfficeXmlNode node) where T : OpenXmlElement {
            if (node.IsNone) return null;
            if (node.IsSdk) return node.SdkElement as T;
            if (presentation == null) {
                throw new InvalidOperationException("A node of the package view needs its live document.");
            }
            _ = presentation.OpenXmlDocument;
            return ToLiveNode(presentation, partName, node).SdkElement as T;
        }

        // SDK callee in LegacyPptWriter.TextProperties.cs.
        private static bool TryWriteMasterParagraphException(Stream output,
            LegacyPptWriterTextStyleSource source, OfficeXmlNode properties,
            LegacyPptWriterFontCatalog fonts, out string? reason) =>
            TryWriteParagraphException(output,
                LiveSdkElement<A.TextParagraphPropertiesType>(source.Presentation, source.PartName, properties),
                fonts, out reason, allowAutoNumbering: true);

        // SDK callee in LegacyPptWriter.TextProperties.cs.
        private static bool TryWriteMasterCharacterException(Stream output,
            LegacyPptWriterTextStyleSource source, OfficeXmlNode properties,
            LegacyPptWriterFontCatalog fonts, out string? reason) =>
            TryWriteCharacterException(output,
                LiveSdkElement<A.DefaultRunProperties>(source.Presentation, source.PartName, properties),
                fonts, out reason);

        // SDK callee in LegacyPptWriter.AutoNumbering.cs; the picture-bullet catalog is keyed by live SDK elements.
        private static bool TryWriteMasterAutomaticNumberingException9(Stream output,
            LegacyPptWriterTextStyleSource source, OfficeXmlNode numbering, int numberingStart,
            OfficeXmlNode pictureBullet,
            LegacyPptWriterPictureBulletCatalog pictureBullets, out string? reason) =>
            TryWriteAutomaticNumberingException9(output,
                LiveSdkElement<A.AutoNumberedBullet>(source.Presentation, source.PartName, numbering),
                numberingStart,
                LiveSdkElement<A.PictureBullet>(source.Presentation, source.PartName, pictureBullet),
                pictureBullets, out reason);

        // The picture catalog in LegacyPptWriter.Pictures.cs is keyed by live SDK elements.
        private static LegacyPptWriterPicture GetBackgroundPicture(
            LegacyPptWriterPictureCatalog catalog, LegacyPptWriterBackground background) =>
            catalog.Get(background.PictureFill
                ?? throw new InvalidOperationException("The picture background fill is missing from the live document."));

        internal sealed partial class LegacyPptWriterBackground {
            // SDK callers in LegacyPptWriter.Pictures.cs and LegacyPptPreservingWriter.Pictures.cs.
            internal A.BlipFill? PictureFill => PictureFillNode.IsNone
                ? null
                : LiveSdkElement<A.BlipFill>(PictureDocument, PicturePartName ?? string.Empty, PictureFillNode);
        }

        // ---- SDK-typed entry points for callers not yet passing a view ----------------------------------------

        private static LegacyPptPackageView CaptureViewOf(OpenXmlPart part, out PowerPointDocument presentation) {
            presentation = PowerPointPartRef.FromSdk(part).Document
                ?? throw new InvalidOperationException("The part is not part of an OfficeIMO presentation.");
            return LegacyPptPackageView.Capture(presentation);
        }

        private static bool TryReadBackgroundOfSdkPart(OpenXmlPart part,
            out LegacyPptWriterBackground? background, out string? reason) {
            ArgumentNullException.ThrowIfNull(part);
            using LegacyPptPackageView view = CaptureViewOf(part, out PowerPointDocument presentation);
            return TryReadBackground(presentation, view.Part(part.Uri.ToString()),
                out background, out reason);
        }

        // SDK callers in LegacyPptWritePreflight.cs, LegacyPptPreservingWriter.Masters.cs and LegacyPptWriter.Pictures.cs.
        internal static bool TryReadBackground(SlideMasterPart masterPart,
            out LegacyPptWriterBackground? background, out string? reason) =>
            TryReadBackgroundOfSdkPart(masterPart, out background, out reason);

        // SDK callers in LegacyPptWritePreflight.cs, LegacyPptPreservingWriter.Masters.cs and LegacyPptWriter.Pictures.cs.
        internal static bool TryReadBackground(NotesMasterPart masterPart,
            out LegacyPptWriterBackground? background, out string? reason) =>
            TryReadBackgroundOfSdkPart(masterPart, out background, out reason);

        // SDK callers in LegacyPptWritePreflight.cs, LegacyPptPreservingWriter.Masters.cs and LegacyPptWriter.Pictures.cs.
        internal static bool TryReadBackground(HandoutMasterPart masterPart,
            out LegacyPptWriterBackground? background, out string? reason) =>
            TryReadBackgroundOfSdkPart(masterPart, out background, out reason);

        // SDK caller in LegacyPptPreservingWriter.Masters.cs.
        internal static bool TryReadBackground(SlideLayoutPart masterPart,
            out LegacyPptWriterBackground? background, out string? reason) =>
            TryReadBackgroundOfSdkPart(masterPart, out background, out reason);

        // SDK callers in LegacyPptPreservingWriter.ChangePlan.cs and LegacyPptWriter.Pictures.cs.
        internal static bool TryReadBackground(NotesSlidePart notesPart,
            out LegacyPptWriterBackground? background, out string? reason) =>
            TryReadBackgroundOfSdkPart(notesPart, out background, out reason);

        private static IReadOnlyList<PowerPointShape> ReadMasterShapesOfSdkPart(
            OpenXmlPart part, out string? unsupportedReason) {
            ArgumentNullException.ThrowIfNull(part);
            using LegacyPptPackageView view = CaptureViewOf(part, out PowerPointDocument presentation);
            return ReadMasterShapesForWrite(presentation, view.Part(part.Uri.ToString()),
                out unsupportedReason);
        }

        // SDK callers in LegacyPptWritePreflight.cs, LegacyPptPreservingWriter.Masters.cs, LegacyPptProjectionMap.cs
        // and LegacyPptWriter.Pictures.cs.
        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            SlideMasterPart masterPart, out string? unsupportedReason) =>
            ReadMasterShapesOfSdkPart(masterPart, out unsupportedReason);

        // SDK callers in LegacyPptWritePreflight.cs, LegacyPptPreservingWriter.Masters.cs, LegacyPptProjectionMap.cs
        // and LegacyPptWriter.Pictures.cs.
        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            NotesMasterPart masterPart, out string? unsupportedReason) =>
            ReadMasterShapesOfSdkPart(masterPart, out unsupportedReason);

        // SDK callers in LegacyPptWritePreflight.cs, LegacyPptPreservingWriter.Masters.cs, LegacyPptProjectionMap.cs
        // and LegacyPptWriter.Pictures.cs.
        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            HandoutMasterPart masterPart, out string? unsupportedReason) =>
            ReadMasterShapesOfSdkPart(masterPart, out unsupportedReason);

        // SDK callers in LegacyPptPreservingWriter.Masters.cs and LegacyPptProjectionMap.cs.
        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            SlideLayoutPart masterPart, out string? unsupportedReason) =>
            ReadMasterShapesOfSdkPart(masterPart, out unsupportedReason);

        private static byte[] BuildPreservedMasterThemeRecordOfSdkPart(
            LegacyPptRecord prototype, OpenXmlPart source,
            IReadOnlyList<int> changedClassicColorSlots) {
            ArgumentNullException.ThrowIfNull(prototype);
            ArgumentNullException.ThrowIfNull(source);
            using LegacyPptPackageView view = CaptureViewOf(source, out _);
            return BuildPreservedMasterThemeRecord(prototype, view.Part(source.Uri.ToString()),
                changedClassicColorSlots);
        }

        // SDK caller in LegacyPptPreservingWriter.Masters.cs.
        internal static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, SlideMasterPart source,
            IReadOnlyList<int> changedClassicColorSlots) =>
            BuildPreservedMasterThemeRecordOfSdkPart(prototype, source, changedClassicColorSlots);

        // SDK caller in LegacyPptPreservingWriter.Masters.cs.
        internal static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, NotesMasterPart source,
            IReadOnlyList<int> changedClassicColorSlots) =>
            BuildPreservedMasterThemeRecordOfSdkPart(prototype, source, changedClassicColorSlots);

        // SDK caller in LegacyPptPreservingWriter.Masters.cs.
        internal static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, HandoutMasterPart source,
            IReadOnlyList<int> changedClassicColorSlots) =>
            BuildPreservedMasterThemeRecordOfSdkPart(prototype, source, changedClassicColorSlots);

        // SDK caller in LegacyPptPreservingWriter.Masters.cs.
        internal static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, SlideLayoutPart source,
            IReadOnlyList<int> changedClassicColorSlots) =>
            BuildPreservedMasterThemeRecordOfSdkPart(prototype, source, changedClassicColorSlots);

        private static byte[] BuildPreservedThemeRecordOfSdkPart(
            LegacyPptRecord prototype, OpenXmlPart source,
            IReadOnlyList<int> changedClassicColorSlots) {
            ArgumentNullException.ThrowIfNull(prototype);
            ArgumentNullException.ThrowIfNull(source);
            using LegacyPptPackageView view = CaptureViewOf(source, out _);
            return BuildPreservedThemeRecord(prototype, view.Part(source.Uri.ToString()),
                changedClassicColorSlots);
        }

        // SDK caller in LegacyPptPreservingWriter.ChangePlan.cs.
        internal static byte[] BuildPreservedThemeRecord(
            LegacyPptRecord prototype, SlidePart source,
            IReadOnlyList<int> changedClassicColorSlots) =>
            BuildPreservedThemeRecordOfSdkPart(prototype, source, changedClassicColorSlots);

        // SDK caller in LegacyPptPreservingWriter.ChangePlan.cs.
        internal static byte[] BuildPreservedThemeRecord(
            LegacyPptRecord prototype, NotesSlidePart source,
            IReadOnlyList<int> changedClassicColorSlots) =>
            BuildPreservedThemeRecordOfSdkPart(prototype, source, changedClassicColorSlots);

        // SDK caller in LegacyPptPreservingWriter.Masters.cs.
        internal static bool TryRewriteMasterTextStyleRecords(
            LegacyPptRecord master, P.TextStyles? styles,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out byte[] bytes,
            out string? reason) {
            OpenXmlPart? part = styles?.Ancestors<OpenXmlPartRootElement>().FirstOrDefault()?.OpenXmlPart;
            if (styles == null || part == null) {
                return TryRewriteMasterTextStyleRecords(master, null, string.Empty, default,
                    fonts, pictureBullets, out bytes, out reason);
            }
            using LegacyPptPackageView view = CaptureViewOf(part, out PowerPointDocument presentation);
            string partName = part.Uri.ToString();
            return TryRewriteMasterTextStyleRecords(master, presentation, partName,
                view.Resolve(partName, OfficeXmlNode.FromSdk(styles)),
                fonts, pictureBullets, out bytes, out reason);
        }

        // SDK caller in LegacyPptWriter.Pictures.cs, which collects the pictures as live SDK elements.
        private static bool TryValidateMaterializedLayoutPictures(
            PowerPointDocument presentation,
            ISet<OpenXmlElement> materializedPictures,
            out string? reason) {
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(presentation);
            var resolved = new HashSet<OfficeXmlNode>();
            foreach (OpenXmlElement picture in materializedPictures) {
                OpenXmlPart? part = picture.Ancestors<OpenXmlPartRootElement>().FirstOrDefault()?.OpenXmlPart;
                if (part == null) continue;
                OfficeXmlNode node = view.Resolve(part.Uri.ToString(), OfficeXmlNode.FromSdk(picture));
                if (!node.IsNone) resolved.Add(node);
            }
            return TryValidateMaterializedLayoutPictures(view, resolved, out reason);
        }

        // SDK caller in PowerPointFeatureReport.cs.
        internal static bool HasCanonicalClassicAuthorColorIndex(
            P.CommentAuthor author) =>
            HasCanonicalClassicAuthorColorIndex(OfficeXmlNode.FromSdk(author));

        // SDK caller in PowerPointFeatureReport.cs.
        internal static bool TryGetBinaryCompatibleClassicAuthorIdentity(
            P.CommentAuthor author,
            out string name,
            out string initials,
            out string identity) =>
            TryGetBinaryCompatibleClassicAuthorIdentity(OfficeXmlNode.FromSdk(author),
                out name, out initials, out identity);
    }
}
