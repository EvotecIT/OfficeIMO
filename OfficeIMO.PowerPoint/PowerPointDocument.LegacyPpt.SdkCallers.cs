using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The binary import's bridge to the table and notes projections, which still build SDK elements
    /// (PowerPointDocument.LegacyPpt.Tables.cs and PowerPointDocument.LegacyPpt.Notes.cs). Each member converts between
    /// SDK parts and elements and the import's <see cref="PowerPointPartRef"/> and <see cref="OfficeXmlNode"/> forms.
    /// Remove a member when its callers are ported.
    /// </summary>
    public sealed partial class PowerPointDocument {
        // Opens the SDK for the scratch presentation, for projections that build through it.
        private static void EnsureLegacySdk(PowerPointDocument document) => document.EnsureSdk();

        // The scratch presentation saved as the engine saves a presentation, VBA project kept. The save entry point's
        // signature names the SDK document type, so the call lives here.
        private static byte[] CreateLegacyProjectedPackageBytes(PowerPointDocument projected) =>
            projected.CreatePackageBytesForSave(options: null, preserveVbaProject: true);

        // SDK caller in PowerPointDocument.LegacyPpt.Tables.cs (deferred table cell interactions)
        internal sealed partial class LegacyPptDeferredProjection {
            internal LegacyPptDeferredProjection(
                Action<IReadOnlyDictionary<uint, SlidePart>> apply)
                : this(projectedSlides => apply(ToSdkSlideParts(projectedSlides)!)) {
                ArgumentNullException.ThrowIfNull(apply);
            }
        }

        private static IReadOnlyDictionary<uint, SlidePart>? ToSdkSlideParts(
            IReadOnlyDictionary<uint, PowerPointPartRef>? slideParts) {
            if (slideParts == null) return null;
            var result = new Dictionary<uint, SlidePart>();
            foreach (KeyValuePair<uint, PowerPointPartRef> item in slideParts) {
                PowerPointPartRef current = item.Value.Current;
                if (current.SdkPart == null && current.Document is { } document) {
                    EnsureLegacySdk(document);
                    current = item.Value.Current;
                }
                result.Add(item.Key, (SlidePart)(current.SdkPart
                    ?? throw new InvalidOperationException("The projected slide is not open in the SDK.")));
            }
            return result;
        }

        private static IReadOnlyDictionary<uint, PowerPointPartRef>? FromSdkSlideParts(
            IReadOnlyDictionary<uint, SlidePart>? slideParts) =>
            slideParts?.ToDictionary(item => item.Key, item => PowerPointPartRef.FromSdk(item.Value));

        // SDK caller in PowerPointDocument.LegacyPpt.Tables.cs
        private static bool ShouldDeferLegacyTextInteractions(
            LegacyPptTextBody textBody,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId,
            ICollection<LegacyPptDeferredProjection>? deferredInteractions) =>
            ShouldDeferLegacyTextInteractions(textBody,
                slidePartsByLegacyId == null ? null : new Dictionary<uint, PowerPointPartRef>(),
                deferredInteractions);

        // SDK caller in PowerPointDocument.LegacyPpt.Tables.cs
        private static IReadOnlyList<OpenXmlElement> ProjectLegacyInteraction(
            OpenXmlPart ownerPart, LegacyPptInteraction interaction,
            bool shapeLevel = false,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId = null,
            LegacyPptSoundProjectionContext? soundContext = null) {
            var container = new A.RunProperties();
            ProjectLegacyInteraction(PowerPointPartRef.FromSdk(ownerPart), interaction,
                OfficeXmlNode.FromSdk(container), shapeLevel,
                FromSdkSlideParts(slidePartsByLegacyId), soundContext);
            OpenXmlElement[] elements = container.ChildElements.ToArray();
            foreach (OpenXmlElement element in elements) element.Remove();
            return elements;
        }

        // SDK caller in PowerPointDocument.LegacyPpt.Tables.cs
        private static string? ProjectLegacyPictureBullet(
            OpenXmlPart ownerPart, LegacyPptPictureBullet pictureBullet) =>
            ProjectLegacyPictureBullet(PowerPointPartRef.FromSdk(ownerPart), pictureBullet);

        // Builds a native table frame through the table projection and returns it detached.
        private static OfficeXmlNode CreateLegacyTableFrameElement(
            PowerPointPartRef ownerPart, LegacyPptShape source, uint shapeId,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext,
            ICollection<LegacyPptDeferredProjection>? deferredInteractions) {
            PowerPointPartRef current = ownerPart.Current;
            if (current.SdkPart == null && current.Document is { } document) {
                EnsureLegacySdk(document);
                current = ownerPart.Current;
            }
            var part = (OpenXmlPart)(current.SdkPart
                ?? throw new InvalidOperationException("The projected part is not open in the SDK."));
            return OfficeXmlNode.FromSdk(CreateLegacyTableFrame(part, source, shapeId,
                ToSdkSlideParts(slidePartsByLegacyId), soundContext, deferredInteractions));
        }

        // Projects a slide table through the table projection.
        private static PowerPointShape ProjectLegacyTableThroughSdk(PowerPointSlide slide,
            LegacyPptShape source,
            IReadOnlyDictionary<uint, PowerPointPartRef> slidePartsByLegacyId,
            LegacyPptSoundProjectionContext soundContext) =>
            ProjectLegacyTable(slide, source, ToSdkSlideParts(slidePartsByLegacyId)!, soundContext);

        // Projects a notes page through the notes projection.
        private static void ProjectLegacyNotesPageThroughSdk(PowerPointSlide slide,
            LegacyPptNotesPage source,
            IReadOnlyDictionary<uint, PowerPointPartRef> slidePartsByLegacyId,
            LegacyPptSoundProjectionContext soundContext) =>
            ProjectLegacyNotesPage(slide, source, ToSdkSlideParts(slidePartsByLegacyId)!, soundContext);

        // SDK caller in PowerPointDocument.LegacyPpt.Notes.cs
        private static P.ShapeTree CreateLegacyShapeTree(OpenXmlPart ownerPart,
            IReadOnlyList<LegacyPptShape>? shapes,
            IReadOnlyList<LegacyPptConnectorRule>? connectorRules,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext) {
            var tree = new P.ShapeTree();
            WriteLegacyShapeTree(OfficeXmlNode.FromSdk(tree), PowerPointPartRef.FromSdk(ownerPart), shapes,
                connectorRules, FromSdkSlideParts(slidePartsByLegacyId), soundContext);
            return tree;
        }

        // SDK caller in PowerPointDocument.LegacyPpt.Notes.cs
        private static void ApplyLegacyRoundTripTheme(OpenXmlPart target, LegacyPptRoundTripTheme? source) {
            PowerPointPartRef part = PowerPointPartRef.FromSdk(target);
            ApplyLegacyRoundTripTheme(LegacyPptImportParts.For(RequireLegacyDocument(part)), part, source);
        }

        // SDK caller in PowerPointDocument.LegacyPpt.Notes.cs
        private static void ApplyLegacyBackground(OpenXmlPart ownerPart,
            P.CommonSlideData commonSlideData, LegacyPptBackground source) =>
            ApplyLegacyBackground(PowerPointPartRef.FromSdk(ownerPart), OfficeXmlNode.FromSdk(commonSlideData), source);

        // SDK caller in PowerPointDocument.LegacyPpt.Notes.cs
        private static void ApplyLegacyBackground(P.CommonSlideData commonSlideData, string color) =>
            ApplyLegacyBackground(OfficeXmlNode.FromSdk(commonSlideData), color);

        // SDK caller in PowerPointDocument.LegacyPpt.Notes.cs
        private static void SetLegacyThemeColors(A.ColorScheme target, LegacyPptColorScheme source) =>
            SetLegacyThemeColors(OfficeXmlNode.FromSdk(target), source);
    }
}
