using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    internal static partial class WordFieldInventory {
        // Existing SDK consumers, including PDF field diagnostics, keep their original element identity.
        internal static IReadOnlyList<FieldInfoCandidate> InspectWithSource(WordDocument document) =>
            InspectStories(WordPackageView.ForSdkPackage(document._wordprocessingDocument).Main is { } main
                ? EnumerateFieldRoots(main) : Enumerable.Empty<NodeFieldRoot>());

        internal static OfficeXmlNode BodyNode(WordDocument document) =>
            document._ownedModel is { } model
                ? new OfficeXmlNode(model.Document, model.Body)
                : OfficeXmlNode.FromSdk(document.TryGetBodyRoot());

        internal static IEnumerable<FieldRoot> EnumerateFieldRoots(MainDocumentPart mainPart) {
            WordPartView? main = WordPackageView.ForSdkPackage(mainPart.OpenXmlPackage).Main;
            if (main == null) yield break;
            foreach (NodeFieldRoot root in EnumerateFieldRoots(main)) {
                yield return new FieldRoot((OpenXmlCompositeElement)root.Root.SdkElement!, root.LocationKind, root.PartUri);
            }
        }

        internal sealed class FieldRoot {
            internal FieldRoot(OpenXmlCompositeElement root, WordFieldLocationKind locationKind, string partUri) {
                Root = root;
                LocationKind = locationKind;
                PartUri = partUri;
            }

            internal OpenXmlCompositeElement Root { get; }

            internal WordFieldLocationKind LocationKind { get; }

            internal string PartUri { get; }
        }

        internal sealed partial class FieldInfoCandidate {
            internal OpenXmlElement SourceElement => (OpenXmlElement)_source.SdkElement!;
        }
    }
}
