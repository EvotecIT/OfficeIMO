using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    // Temporary bridge: the drawing, picture and preflight readers read the package view, but these members of the
    // master-shape, background and picture-bullet readers still take SDK parts. Each shim finds the SDK part with the
    // view part's name and calls the existing member. Replace each call with the member's view overload once that
    // reader reads the view, then delete this file.
    internal static partial class LegacyPptWriter {
        internal static IReadOnlyList<PowerPointShape> SdkReadMasterShapesForWrite(
            PowerPointDocument presentation, LegacyPptPartView part,
            out string? unsupportedReason) {
            switch (FindSdkPart(presentation, part)) {
                case SlideMasterPart master:
                    return ReadMasterShapesForWrite(master, out unsupportedReason);
                case NotesMasterPart notesMaster:
                    return ReadMasterShapesForWrite(notesMaster, out unsupportedReason);
                case HandoutMasterPart handoutMaster:
                    return ReadMasterShapesForWrite(handoutMaster, out unsupportedReason);
                case SlideLayoutPart layout:
                    return ReadMasterShapesForWrite(layout, out unsupportedReason);
                default:
                    throw new InvalidOperationException($"Part '{part.Name}' is not a master or layout part.");
            }
        }

        internal static bool SdkTryReadBackground(PowerPointDocument presentation,
            LegacyPptPartView part, out LegacyPptWriterBackground? background,
            out string? reason) {
            switch (FindSdkPart(presentation, part)) {
                case SlideMasterPart master:
                    return TryReadBackground(master, out background, out reason);
                case NotesMasterPart notesMaster:
                    return TryReadBackground(notesMaster, out background, out reason);
                case HandoutMasterPart handoutMaster:
                    return TryReadBackground(handoutMaster, out background, out reason);
                case SlideLayoutPart layout:
                    return TryReadBackground(layout, out background, out reason);
                case NotesSlidePart notes:
                    return TryReadBackground(notes, out background, out reason);
                default:
                    throw new InvalidOperationException($"Part '{part.Name}' has no background reader.");
            }
        }

        /// <summary>The background's picture fill element, the key its picture has in the picture catalog.</summary>
        internal static OfficeXmlNode SdkBackgroundPictureFill(LegacyPptWriterBackground background) =>
            OfficeXmlNode.FromSdk(background.PictureFill);

        internal static bool SdkTryReadPictureBulletCatalog(PowerPointSlide slide,
            out LegacyPptWriterPictureBulletCatalog catalog, out string? reason) {
            PresentationPart? presentationPart = slide.SlidePart
                .GetParentParts().OfType<PresentationPart>().FirstOrDefault();
            if (presentationPart == null) {
                catalog = LegacyPptWriterPictureBulletCatalog.Empty;
                reason = null;
                return false;
            }
            return TryReadPictureBulletCatalog(presentationPart, out catalog, out reason);
        }

        private static OpenXmlPart? FindSdkPart(PowerPointDocument presentation, LegacyPptPartView part) {
            PresentationPart? presentationPart = presentation.OpenXmlDocument.PresentationPart;
            if (presentationPart == null) return null;
            var visited = new HashSet<OpenXmlPart>();
            var pending = new Stack<OpenXmlPart>();
            pending.Push(presentationPart);
            while (pending.Count > 0) {
                OpenXmlPart candidate = pending.Pop();
                if (!visited.Add(candidate)) continue;
                if (string.Equals(candidate.Uri.ToString(), part.Name, StringComparison.OrdinalIgnoreCase)) return candidate;
                foreach (IdPartPair child in candidate.Parts) pending.Push(child.OpenXmlPart);
            }
            return null;
        }
    }
}
