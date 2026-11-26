using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The SDK side of a text run: constructors over SDK elements and SDK views of the run, for code that still works
    /// on SDK elements. Reaching an SDK view of a run held in a package tree opens the SDK.
    /// </summary>
    public partial class PowerPointTextRun {
        internal PowerPointTextRun(A.Run run, SlidePart? slidePart = null, OpenXmlPartContainer? ownerPart = null)
            : this(PowerPointXmlNode.FromSdk(run), PowerPointPartRef.FromSdk(slidePart ?? ownerPart as SlidePart), PowerPointPartRef.FromSdk(ownerPart as OpenXmlPart ?? slidePart)) {
        }

        internal PowerPointTextRun(A.Run run, SlidePart? slidePart, OpenXmlPartContainer? ownerPart, List<OpenXmlElement> lineContinuation)
            : this(PowerPointXmlNode.FromSdk(run), PowerPointPartRef.FromSdk(slidePart ?? ownerPart as SlidePart), PowerPointPartRef.FromSdk(ownerPart as OpenXmlPart ?? slidePart),
                lineContinuation.Select(PowerPointXmlNode.FromSdk).ToList()) {
        }

        internal PowerPointTextRun(A.Field field, SlidePart? slidePart = null, OpenXmlPartContainer? ownerPart = null)
            : this(PowerPointXmlNode.FromSdk(field), PowerPointPartRef.FromSdk(slidePart ?? ownerPart as SlidePart), PowerPointPartRef.FromSdk(ownerPart as OpenXmlPart ?? slidePart)) {
        }

        private A.Run? _detachedFieldRun;

        /// <summary>The SDK run; a field has none of its own, so it reads as a detached run.</summary>
        internal A.Run Run => SdkElement() as A.Run ?? (_detachedFieldRun ??= new A.Run());

        /// <summary>The SDK run properties of the run or field.</summary>
        internal A.RunProperties? RunProperties => SdkElement() switch {
            A.Field textField => textField.RunProperties,
            A.Run run => run.RunProperties,
            _ => null
        };

        internal A.RunProperties EnsureRunProperties() {
            A.RunProperties properties = (A.RunProperties)EnsurePropertiesSdkElement();
            return properties;
        }

        /// <summary>The last SDK run of the text this wrapper wrote.</summary>
        internal A.Run LastLine => (A.Run)ToSdk(LastLineElement);

        /// <summary>The SDK part that holds the run.</summary>
        internal OpenXmlPartContainer? OwnerPart => ToSdkPart(OwnerPartRef);

        /// <summary>The SDK part that holds the run, falling back to its slide part.</summary>
        internal OpenXmlPart? SdkOwnerPart => ToSdkPart(OwnerPartRef) ?? ToSdkPart(SlidePartRef);

        private OpenXmlElement SdkElement() => ToSdk(Element);

        private OpenXmlElement EnsurePropertiesSdkElement() => ToSdk(EnsurePropertiesElement());

        // An element as an SDK element, opening the SDK for one held in a package tree.
        private OpenXmlElement ToSdk(PowerPointXmlNode node) {
            if (node.SdkElement is OpenXmlElement element) return element;
            if (_packageSlide is { RemovedFromPackage: true }) throw new InvalidOperationException("The slide is no longer part of the presentation.");
            if (node.Tree != null && !_packagePart.IsNone) return (OpenXmlElement)PowerPointSlide.BindPackageNode(_packagePart, node.Tree, node.Node);
            throw new InvalidOperationException("The run is not held in an SDK part.");
        }

        private static OpenXmlPart? ToSdkPart(PowerPointPartRef part) {
            if (part.IsNone) return null;
            if (part.Current.SdkPart == null && part.Document is { } document) _ = document.OpenXmlDocument;
            return part.Current.SdkPart as OpenXmlPart;
        }

        /// <summary>Removes a content subtree and the relationships no surviving content uses.</summary>
        internal static void DiscardContent(OpenXmlElement content, OpenXmlPart? ownerPart) =>
            DiscardContent(PowerPointXmlNode.FromSdk(content), PowerPointPartRef.FromSdk(ownerPart));
    }
}
