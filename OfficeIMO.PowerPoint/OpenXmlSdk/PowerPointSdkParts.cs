using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint.OpenXmlSdk {
    /// <summary>
    /// The SDK side of <see cref="PowerPointPartRef"/>: part names, roots and relationships of SDK parts, through the
    /// SDK's own calls, so a caller holding the SDK package sees the relationships it would have made itself. Parts are
    /// passed as <see cref="object"/>.
    /// </summary>
    internal static class PowerPointSdkParts {
        private static OpenXmlPart P(object part) => (OpenXmlPart)part;

        internal static PowerPointDocument? DocumentOf(object part) => P(part).OpenXmlPackage.Features.Get<PowerPointDocument>();

        internal static string Name(object part) => P(part).Uri.OriginalString;

        internal static object? Root(object part) => P(part).RootElement;

        internal static PowerPointPartKind Kind(object part) => P(part) switch {
            SlidePart => PowerPointPartKind.Slide,
            SlideLayoutPart => PowerPointPartKind.SlideLayout,
            SlideMasterPart => PowerPointPartKind.SlideMaster,
            NotesSlidePart => PowerPointPartKind.NotesSlide,
            NotesMasterPart => PowerPointPartKind.NotesMaster,
            _ => PowerPointPartKind.Other
        };

        internal static bool TryGetRelationship(object part, string id, out string type, out string target, out bool external, out object? targetPart, out Uri? uri) {
            OpenXmlPart owner = P(part);
            type = target = string.Empty;
            external = false;
            targetPart = null;
            uri = null;
            foreach (HyperlinkRelationship hyperlink in owner.HyperlinkRelationships) {
                if (!string.Equals(hyperlink.Id, id, StringComparison.Ordinal)) continue;
                uri = hyperlink.Uri;
                type = hyperlink.RelationshipType;
                target = hyperlink.Uri.OriginalString;
                external = hyperlink.IsExternal;
                if (!external) targetPart = ResolveRelativePart(owner, hyperlink.Uri);
                return true;
            }
            foreach (ExternalRelationship relationship in owner.ExternalRelationships) {
                if (!string.Equals(relationship.Id, id, StringComparison.Ordinal)) continue;
                uri = relationship.Uri;
                type = relationship.RelationshipType;
                target = relationship.Uri.OriginalString;
                external = true;
                return true;
            }
            foreach (DataPartReferenceRelationship relationship in owner.DataPartReferenceRelationships) {
                if (!string.Equals(relationship.Id, id, StringComparison.Ordinal)) continue;
                type = relationship.RelationshipType;
                target = relationship.Uri.OriginalString;
                return true;
            }
            if (owner.TryGetPartById(id, out OpenXmlPart? related)) {
                type = related.RelationshipType;
                target = related.Uri.OriginalString;
                targetPart = related;
                return true;
            }
            return false;
        }

        // An internal hyperlink names its target relative to the source; the SDK part with that name, if there is one.
        private static OpenXmlPart? ResolveRelativePart(OpenXmlPart owner, Uri relative) {
            Uri root = new("http://officeimo.invalid/", UriKind.Absolute);
            Uri resolved = new(new Uri(new Uri(root, owner.Uri.OriginalString.TrimStart('/')), "."), relative);
            string name = "/" + root.MakeRelativeUri(resolved);
            return owner.OpenXmlPackage.GetAllParts().FirstOrDefault(candidate => string.Equals(candidate.Uri.OriginalString, name, StringComparison.OrdinalIgnoreCase));
        }

        internal static IEnumerable<object> RelatedParts(object part, string relationshipType) {
            foreach (IdPartPair pair in P(part).Parts) {
                if (string.Equals(pair.OpenXmlPart.RelationshipType, relationshipType, StringComparison.Ordinal)) yield return pair.OpenXmlPart;
            }
        }

        internal static string? IdOfPart(object part, object target) => P(part).Parts.FirstOrDefault(pair => ReferenceEquals(pair.OpenXmlPart, target)).RelationshipId;

        internal static string AddHyperlink(object part, Uri uri, bool external) => P(part).AddHyperlinkRelationship(uri, external).Id;

        internal static string GetOrAddRelativeHyperlink(object part, object target) {
            OpenXmlPart owner = P(part);
            Uri targetUri = PowerPointHyperlinkResolver.CreatePartRelativeUri(owner.Uri.OriginalString, P(target).Uri.OriginalString);
            HyperlinkRelationship relationship = owner.HyperlinkRelationships.FirstOrDefault(candidate => !candidate.IsExternal && candidate.Uri == targetUri)
                ?? owner.AddHyperlinkRelationship(targetUri, false);
            return relationship.Id;
        }

        internal static void ReplaceRelatedPart(object part, string relationshipType, object target) {
            OpenXmlPart owner = P(part);
            foreach (IdPartPair pair in owner.Parts.Where(pair => pair.OpenXmlPart.RelationshipType == relationshipType).ToList()) owner.DeletePart(pair.RelationshipId);
            owner.AddPart(P(target));
        }

        internal static string GetOrAddPartRelationship(object part, object target) {
            OpenXmlPart owner = P(part), related = P(target);
            if (!owner.Parts.Any(pair => ReferenceEquals(pair.OpenXmlPart, related))) owner.AddPart(related);
            return owner.GetIdOfPart(related);
        }

        internal static bool ReferencesRelationship(object element, string id) {
            OpenXmlElement root = (OpenXmlElement)element;
            if (HasRelationshipAttribute(root, id)) return true;
            foreach (OpenXmlElement descendant in root.Descendants()) {
                if (HasRelationshipAttribute(descendant, id)) return true;
            }
            return false;
        }

        private static bool HasRelationshipAttribute(OpenXmlElement element, string id) {
            if (!element.HasAttributes) return false;
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (string.Equals(attribute.NamespaceUri, PowerPointUtils.RelationshipIdNamespace, StringComparison.Ordinal)
                    && string.Equals(attribute.Value, id, StringComparison.Ordinal)) {
                    return true;
                }
            }
            return false;
        }

        /// <summary>Removes a relationship the part's content no longer uses, with media only it used.</summary>
        internal static void RemoveRelationshipIfUnused(object part, string id, object? owningSlide) {
            OpenXmlPart owner = P(part);
            if (owner.RootElement != null && ReferencesRelationship(owner.RootElement, id)) return;
            if (owner is NotesSlidePart && owningSlide != null
                && owner.Parts.Any(pair => string.Equals(pair.RelationshipId, id, StringComparison.Ordinal) && ReferenceEquals(pair.OpenXmlPart, owningSlide))) return;
            HyperlinkRelationship? hyperlink = owner.HyperlinkRelationships.FirstOrDefault(candidate => string.Equals(candidate.Id, id, StringComparison.Ordinal));
            if (hyperlink != null) {
                owner.DeleteReferenceRelationship(hyperlink);
                return;
            }
            DataPartReferenceRelationship? media = owner.DataPartReferenceRelationships.FirstOrDefault(candidate => string.Equals(candidate.Id, id, StringComparison.Ordinal));
            if (media != null) {
                DataPart data = media.DataPart;
                owner.DeleteReferenceRelationship(media);
                if (data is MediaDataPart mediaPart && !mediaPart.GetDataPartReferenceRelationships().Any() && owner.OpenXmlPackage is PresentationDocument document) {
                    document.DeletePart(mediaPart);
                }
                return;
            }
            if (owner.Parts.Any(pair => string.Equals(pair.RelationshipId, id, StringComparison.Ordinal))) owner.DeletePart(id);
        }
    }
}

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>The slide parts of the SDK's live slide list, in order.</summary>
        internal IReadOnlyList<OpenXml.PowerPointPartRef> SdkSlideOrder() {
            var order = new List<OpenXml.PowerPointPartRef>();
            PresentationPart? presentation = _sdkDocument?.PresentationPart;
            DocumentFormat.OpenXml.Presentation.SlideIdList? list = presentation?.Presentation?.SlideIdList;
            if (presentation == null || list == null) return order;
            foreach (DocumentFormat.OpenXml.Presentation.SlideId slideId in list.Elements<DocumentFormat.OpenXml.Presentation.SlideId>()) {
                if (slideId.RelationshipId?.Value is { } id && presentation.TryGetPartById(id, out OpenXmlPart? part)) order.Add(OpenXml.PowerPointPartRef.FromSdk(part));
            }
            return order;
        }

        /// <summary>The SDK part with this name once the SDK opened the package, or null.</summary>
        internal object? FindSdkPart(string partName) {
            if (_sdkDocument == null) return null;
            foreach (OpenXmlPart part in _sdkDocument.GetAllParts()) {
                if (string.Equals(part.Uri.OriginalString, partName, StringComparison.OrdinalIgnoreCase)) return part;
            }
            return null;
        }
    }
}
