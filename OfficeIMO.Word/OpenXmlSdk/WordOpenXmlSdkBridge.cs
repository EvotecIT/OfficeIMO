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
        /// The SDK element at the same place as <paramref name="node"/> in the owned tree the package was built from, or
        /// null when the SDK tree no longer has one there. Positions count elements only, as both trees hold the same
        /// markup: text between elements, comments and processing instructions are skipped on both sides.
        /// </summary>
        internal static SdkElement? FindElement(SdkElement sdkRoot, OfficeXmlTree tree, int ownedRoot, int node) {
            if (node == OfficeXmlTree.NoNode) return null;
            var path = new List<int>();
            for (int current = node; current != ownedRoot; current = tree.ParentOf(current)) {
                if (current == OfficeXmlTree.NoNode) return null;
                path.Add(ElementIndex(tree, current));
            }
            SdkElement element = sdkRoot;
            for (int step = path.Count - 1; step >= 0; step--) {
                SdkElement? child = ChildAt(element, path[step]);
                if (child == null) return null;
                element = child;
            }
            return element;
        }

        private static int ElementIndex(OfficeXmlTree tree, int node) {
            int index = 0;
            for (int sibling = tree.PreviousSiblingOf(node); sibling != OfficeXmlTree.NoNode; sibling = tree.PreviousSiblingOf(sibling)) {
                if (tree.KindOf(sibling) == OfficeXmlTreeNodeKind.Element) index++;
            }
            return index;
        }

        private static SdkElement? ChildAt(SdkElement parent, int index) {
            for (SdkElement? child = parent.FirstChild; child != null; child = child.NextSibling()) {
                if (child is SdkMiscNode) continue;
                if (index-- == 0) return child;
            }
            return null;
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
                LoadDocument();
                WordChart.InitializeAxisIdSeed(wordDocument);
                WordListStyles.InitializeAbstractNumberId(wordDocument);
            } catch {
                packageStream.Dispose();
                throw;
            }
        }
    }
}
