using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Resolves DrawingML click hyperlinks to external URIs or stable slide fragments.
    /// </summary>
    internal static partial class PowerPointHyperlinkResolver {
        internal const string SlideJumpAction = "ppaction://hlinksldjump";
        private const string NextSlideAction = "ppaction://hlinkshowjump?jump=nextslide";

        /// <summary>
        /// The target of a click hyperlink in <paramref name="owner"/>: its external URI, or <c>#slide-N</c> for a jump to
        /// a slide of the presentation; null when it has neither.
        /// </summary>
        internal static Uri? Resolve(PowerPointPartRef owner, PowerPointPartRef sourceSlide, OfficeXmlNode hyperlink) {
            if (owner.IsNone || hyperlink.IsNone) return null;
            string? action = hyperlink.Attribute("action");
            if (string.Equals(action, NextSlideAction, StringComparison.OrdinalIgnoreCase)) {
                // A final-slide next action has no PDF page target; keep it explicit for omission reporting.
                int position = SlideNumberOf(owner.Document, sourceSlide);
                return position == 0 ? null : SlideFragment(position + 1);
            }

            string? relationshipId = hyperlink.Attribute(OfficeXmlNamespaces.Relationships, "id");
            if (string.IsNullOrWhiteSpace(relationshipId) || !owner.TryGetRelationship(relationshipId!, out PowerPointRelationship relationship)) return null;
            bool hyperlinkRelationship = string.Equals(relationship.Type, PowerPointPartRef.HyperlinkRelationshipType, StringComparison.Ordinal);
            if (hyperlinkRelationship && relationship.IsExternal) {
                return relationship.Uri ?? (Uri.TryCreate(relationship.Target, UriKind.RelativeOrAbsolute, out Uri? external) ? external : new UnparsedTargetUri(relationship.Target));
            }
            if (!string.Equals(action, SlideJumpAction, StringComparison.OrdinalIgnoreCase) || relationship.IsExternal) return null;
            int slideNumber = SlideNumberOf(owner.Document, relationship.TargetPart);
            return slideNumber == 0 ? null : SlideFragment(slideNumber);
        }

        /// <summary>The slide part a <c>#slide-N</c> fragment names in the current slide order, or none.</summary>
        internal static PowerPointPartRef ResolveSlideFragment(PowerPointDocument document, Uri uri) {
            if (!TryParseSlideFragment(uri, out int slideNumber)) return default;
            IReadOnlyList<PowerPointPartRef> slides = SlideOrder(document);
            return slideNumber <= slides.Count ? slides[slideNumber - 1] : default;
        }

        // The slide parts in presentation order: the SDK's live slide list once it is open, so order changes made
        // through it count, and the document's slides otherwise.
        private static IReadOnlyList<PowerPointPartRef> SlideOrder(PowerPointDocument document) =>
            document.IsSdkLoaded ? document.SdkSlideOrder() : document.Slides.Select(slide => slide.PartRef).ToList();

        /// <summary>A part's path relative to the folder of <paramref name="sourcePartName"/>, as a relationship target.</summary>
        internal static Uri CreatePartRelativeUri(string sourcePartName, string targetPartName) {
            Uri packageRoot = new Uri("http://officeimo.invalid/", UriKind.Absolute);
            Uri sourceUri = new Uri(packageRoot, sourcePartName.TrimStart('/'));
            Uri sourceDirectory = new Uri(sourceUri, ".");
            Uri targetUri = new Uri(packageRoot, targetPartName.TrimStart('/'));
            return sourceDirectory.MakeRelativeUri(targetUri);
        }

        internal static bool TryParseSlideFragment(Uri? uri, out int slideNumber) {
            slideNumber = 0;
            if (uri == null || uri.IsAbsoluteUri) return false;
            const string prefix = "#slide-";
            string fragment = uri.OriginalString;
            return fragment.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                && int.TryParse(fragment.Substring(prefix.Length), NumberStyles.None,
                    CultureInfo.InvariantCulture, out slideNumber)
                && slideNumber >= 1;
        }

        internal static string GetSlideDestinationName(int slideNumber) =>
            "powerpoint-slide-" + slideNumber.ToString(CultureInfo.InvariantCulture);

        /// <summary>
        /// A target the runtime can't parse, such as <c>mailto:one@</c>, held under a placeholder as the SDK holds it:
        /// it reads back as the original text.
        /// </summary>
        private sealed class UnparsedTargetUri : Uri {
            internal UnparsedTargetUri(string target) : base("rewritten://" + Guid.NewGuid().ToString("D")) => Target = target;

            internal string Target { get; }

            public override string ToString() => Target;
        }

        /// <summary>The text to write as a relationship target for <paramref name="uri"/>.</summary>
        internal static string TargetText(Uri uri) => uri is UnparsedTargetUri unparsed ? unparsed.Target : uri.OriginalString;

        private static Uri SlideFragment(int slideNumber) =>
            new("#slide-" + slideNumber.ToString(CultureInfo.InvariantCulture), UriKind.Relative);

        // The 1-based position of the slide part in the presentation, or 0 when it is not one of its slides.
        private static int SlideNumberOf(PowerPointDocument? document, PowerPointPartRef slide) {
            if (document == null || slide.IsNone) return 0;
            IReadOnlyList<PowerPointPartRef> slides = SlideOrder(document);
            for (int index = 0; index < slides.Count; index++) {
                if (slides[index].Equals(slide)) return index + 1;
            }
            return 0;
        }
    }
}
