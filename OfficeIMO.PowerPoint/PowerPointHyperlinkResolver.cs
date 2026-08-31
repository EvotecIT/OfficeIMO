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

            HyperlinkRelationship? relationship = ownerPart.HyperlinkRelationships
                .FirstOrDefault(relationship => string.Equals(
                    relationship.Id, relationshipId, StringComparison.Ordinal));
            if (relationship?.IsExternal == true) return relationship.Uri;

            if (!string.Equals(hyperlink.Action?.Value, SlideJumpAction,
                    StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            PresentationPart? presentationPart = sourceSlidePart?
                .GetParentParts().OfType<PresentationPart>().FirstOrDefault();
            SlidePart? targetSlidePart = null;
            if (relationship?.IsExternal == false && ownerPart is OpenXmlPart hyperlinkOwner) {
                Uri targetUri = ResolvePartUri(hyperlinkOwner.Uri, relationship.Uri);
                targetSlidePart = presentationPart?.SlideParts.FirstOrDefault(candidate =>
                    candidate.Uri == targetUri);
            } else if (ownerPart.TryGetPartById(relationshipId!,
                    out OpenXmlPart? targetPart)) {
                targetSlidePart = targetPart as SlidePart;
            }

            P.SlideIdList? slideIds = presentationPart?.Presentation?
                .SlideIdList;
            if (presentationPart == null || slideIds == null || targetSlidePart == null) return null;

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

        internal static Uri CreatePartRelativeUri(OpenXmlPart sourcePart,
            OpenXmlPart targetPart) {
            Uri packageRoot = new Uri("http://officeimo.invalid/", UriKind.Absolute);
            Uri sourceUri = new Uri(packageRoot,
                sourcePart.Uri.OriginalString.TrimStart('/'));
            Uri sourceDirectory = new Uri(sourceUri, ".");
            Uri targetUri = new Uri(packageRoot,
                targetPart.Uri.OriginalString.TrimStart('/'));
            return sourceDirectory.MakeRelativeUri(targetUri);
        }

        internal static bool TryResolveSlideFragment(SlidePart sourceSlidePart,
            Uri uri, out SlidePart? targetSlidePart) {
            targetSlidePart = null;
            if (uri == null || uri.IsAbsoluteUri) return false;

            const string prefix = "#slide-";
            string fragment = uri.OriginalString;
            if (!fragment.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                || !int.TryParse(fragment.Substring(prefix.Length), NumberStyles.None,
                    CultureInfo.InvariantCulture, out int slideNumber)
                || slideNumber < 1) {
                return false;
            }

            PresentationPart? presentationPart = sourceSlidePart.GetParentParts()
                .OfType<PresentationPart>().FirstOrDefault();
            P.SlideIdList? slideIds = presentationPart?.Presentation?.SlideIdList;
            P.SlideId? slideId = slideIds?.Elements<P.SlideId>()
                .Skip(slideNumber - 1).FirstOrDefault();
            string? relationshipId = slideId?.RelationshipId?.Value;
            if (presentationPart == null || string.IsNullOrWhiteSpace(relationshipId)
                || !presentationPart.TryGetPartById(relationshipId!, out OpenXmlPart? targetPart)
                || targetPart is not SlidePart slidePart) {
                return false;
            }

            targetSlidePart = slidePart;
            return true;
        }

        private static Uri ResolvePartUri(Uri sourcePartUri,
            Uri relationshipUri) {
            Uri packageRoot = new Uri("http://officeimo.invalid/", UriKind.Absolute);
            Uri sourceUri = new Uri(packageRoot,
                sourcePartUri.OriginalString.TrimStart('/'));
            Uri sourceDirectory = new Uri(sourceUri, ".");
            Uri resolved = new Uri(sourceDirectory, relationshipUri);
            return new Uri("/" + packageRoot.MakeRelativeUri(resolved), UriKind.Relative);
        }
    }
}
