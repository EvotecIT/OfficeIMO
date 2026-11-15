using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointPackageParts {
        /// <summary>
        /// Returns the presentation part that relates directly to <paramref name="part"/>, as
        /// <c>part.GetParentParts().OfType&lt;PresentationPart&gt;().FirstOrDefault()</c> does.
        /// <see cref="OpenXmlPart.GetParentParts"/> walks every reachable part in the package on
        /// each call, which made per-run and per-cell lookups grow with the slide count. Only the
        /// package's top-level presentation part can be a presentation parent, so this checks its
        /// relationships instead.
        /// </summary>
        internal static PresentationPart? GetParentPresentationPart(this OpenXmlPart? part) {
            if (part == null) return null;
            foreach (IdPartPair root in part.OpenXmlPackage.Parts) {
                if (root.OpenXmlPart is not PresentationPart presentation) continue;
                foreach (IdPartPair child in presentation.Parts) {
                    if (ReferenceEquals(child.OpenXmlPart, part)) return presentation;
                }
            }
            return null;
        }
    }
}
