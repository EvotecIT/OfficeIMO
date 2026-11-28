using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using SdkElement = DocumentFormat.OpenXml.OpenXmlElement;
using SdkMiscNode = DocumentFormat.OpenXml.OpenXmlMiscNode;

namespace OfficeIMO.Word {
    /// <summary>
    /// The one place where a document held by OfficeIMO's own Open XML layer meets the Open XML SDK. Its contract with
    /// the engine takes package bytes and owned tree positions, never SDK types: the engine hands over the document as
    /// it stands, the bridge opens the SDK package over those bytes, and the SDK package is the document from then on.
    /// It is written to move into an opt-in package later, with the SDK dependency.
    /// </summary>
    internal static class WordOpenXmlSdkBridge {
        /// <summary>Opens the SDK package for <paramref name="document"/> over <paramref name="packageBytes"/>.</summary>
        internal static void OpenPackage(WordDocument document, byte[] packageBytes, WordSdkDeferredLoad load) =>
            document.AttachSdkPackage(packageBytes, load);

        /// <summary>
        /// Pairs every element of the owned main document tree with the SDK element built from it, by walking both trees
        /// together right after the SDK package opens, before anything edits it. Paragraph objects handed out earlier
        /// then find their elements whatever later edits move. Both trees hold the same markup; text between elements,
        /// comments and processing instructions are skipped on both sides.
        /// </summary>
        internal static object?[] MapOwnedTree(OfficeXmlTree tree, SdkElement sdkRoot) {
            var map = new object?[tree.NodeCount];
            Map(tree, tree.Root, sdkRoot, map);
            return map;
        }

        private static void Map(OfficeXmlTree tree, int node, SdkElement element, object?[] map) {
            map[node] = element;
            SdkElement? sdkChild = NextElement(element.FirstChild);
            for (int child = tree.FirstChildOf(node); child != OfficeXmlTree.NoNode && sdkChild != null; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                Map(tree, child, sdkChild, map);
                sdkChild = NextElement(sdkChild.NextSibling());
            }
        }

        private static SdkElement? NextElement(SdkElement? element) {
            while (element is SdkMiscNode) element = element.NextSibling();
            return element;
        }
    }

    public partial class WordDocument {
        /// <summary>
        /// Opens the SDK package over the bytes of a document that was loaded without it, with the load's settings,
        /// and builds the document model on it as a load through the SDK would have.
        /// </summary>
        internal void AttachSdkPackage(byte[] packageBytes, WordSdkDeferredLoad load) {
            OpenSettings settings = CreateOpenSettings(load.OpenSettings);
            Stream packageStream = OfficePackageLoadSource.CreatePackageStream(packageBytes,
                load.ReadOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite);
            try {
                WordZipPackage? zipPackage = null;
                WordprocessingDocument wordDocument = load.ReadOnly
                    ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlReadOnlyPackage.OpenWordprocessing(packageStream, settings)
                    : OpenOwnedEditablePackage(packageStream, settings, out zipPackage);
                _ownedPackageStream = packageStream;
                _ownedZipPackage = zipPackage;
                _wordprocessingDocument = wordDocument;
                _document = wordDocument.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document is missing.");
                if (load.Created) {
                    // As CreateInternal sets up a document over the default package; the identifier state is still the
                    // fresh one it would reset. The owned layer doesn't edit the styles or theme parts, so they are still
                    // the default package's and its cached catalogs apply.
                    InitializeNewDocumentFacade(this);
                    _defaultPackageThemePart = wordDocument.MainDocumentPart?.ThemePart;
                    _defaultPackageStylesPart = wordDocument.MainDocumentPart?.StyleDefinitionsPart;
                    _defaultPackageType = WordprocessingDocumentType.Document;
                } else {
                    LoadDocument();
                    WordChart.InitializeAxisIdSeed(wordDocument);
                }
            } catch {
                packageStream.Dispose();
                _ownedPackageStream = null;
                _ownedZipPackage = null;
                _wordprocessingDocumentValue = null;
                _documentValue = null;
                _defaultPackageThemePart = null;
                _defaultPackageStylesPart = null;
                _sections.Clear();
                throw;
            }
        }
    }
}
