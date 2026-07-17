using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Resolves DrawingML click hyperlinks to external URIs or stable slide fragments.
    /// </summary>
    internal static class PowerPointHyperlinkResolver {
        private const string SlideJumpAction = "ppaction://hlinksldjump";

        internal static Uri? Resolve(OpenXmlPartContainer ownerPart,
            SlidePart? sourceSlidePart, A.HyperlinkOnClick? hyperlink) {
            if (ownerPart == null || hyperlink == null) return null;

            string? relationshipId = hyperlink.Id?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) return null;

            HyperlinkRelationship? external = ownerPart.HyperlinkRelationships
                .FirstOrDefault(relationship => string.Equals(
                    relationship.Id, relationshipId, StringComparison.Ordinal));
            if (external != null) return external.Uri;

            if (!string.Equals(hyperlink.Action?.Value, SlideJumpAction,
                    StringComparison.OrdinalIgnoreCase)
                || !ownerPart.TryGetPartById(relationshipId!,
                    out OpenXmlPart? targetPart)
                || targetPart is not SlidePart targetSlidePart) {
                return null;
            }

            PresentationPart? presentationPart = sourceSlidePart?
                .GetParentParts().OfType<PresentationPart>().FirstOrDefault();
            P.SlideIdList? slideIds = presentationPart?.Presentation?
                .SlideIdList;
            if (presentationPart == null || slideIds == null) return null;

            int slideNumber = 0;
            foreach (P.SlideId slideId in slideIds.Elements<P.SlideId>()) {
                slideNumber++;
                string? targetRelationshipId = slideId.RelationshipId?.Value;
                if (string.IsNullOrWhiteSpace(targetRelationshipId)
                    || !presentationPart.TryGetPartById(targetRelationshipId!,
                        out OpenXmlPart? candidate)
                    || !ReferenceEquals(candidate, targetSlidePart)) {
                    continue;
                }

                return new Uri("#slide-" + slideNumber.ToString(
                    CultureInfo.InvariantCulture), UriKind.Relative);
            }

            return null;
        }
    }
}
