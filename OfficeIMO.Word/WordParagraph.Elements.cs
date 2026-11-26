using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>
    /// The paragraph's elements as <see cref="OfficeXmlNode"/> handles: in OfficeIMO's own tree while the document works
    /// without the Open XML SDK, or the SDK's elements once the SDK holds it. Formatting members read and edit through
    /// these, so each has one implementation for both.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>The paragraph element.</summary>
        private OfficeXmlNode ParagraphNode =>
            OwnedModel != null ? new OfficeXmlNode(_ownedTree!, _ownedParagraphNode) : OfficeXmlNode.FromSdk(SdkParagraphElement);

        /// <summary>Whether the paragraph wraps a run.</summary>
        private bool HasRun => OwnedModel != null ? _ownedRunNode != NoNode : SdkHasRun;

        /// <summary>The properties of the paragraph's own run, never a hyperlink's or a content control's.</summary>
        private OfficeXmlNode OwnRunPropertiesNode {
            get {
                if (OwnedModel != null) return RunPropertiesNode;
                return OfficeXmlNode.FromSdk(SdkOwnRunPropertiesElement);
            }
        }

        /// <summary>The paragraph properties to read, or <see cref="OfficeXmlNode.None"/> when the paragraph has none.</summary>
        private OfficeXmlNode ParagraphPropertiesNode {
            get {
                if (OwnedModel != null) return new OfficeXmlNode(_ownedTree!, _ownedParagraphNode).Child(W, "pPr");
                return OfficeXmlNode.FromSdk(SdkParagraphPropertiesElement);
            }
        }

        /// <summary>The paragraph properties to edit, created as the paragraph's first child when missing.</summary>
        private OfficeXmlNode EditParagraphProperties() {
            WordPackageModel? model = OwnedModel;
            if (model == null) return OfficeXmlNode.FromSdk(SdkEditParagraphPropertiesElement());
            OfficeXmlTree tree = _ownedTree!;
            WordOwnedEdits.PrepareBodyEdit(model);
            int properties = tree.GetOrInsertChild(_ownedParagraphNode, W, "pPr", WPView.Order);
            model.MarkChanged();
            return new OfficeXmlNode(tree, properties);
        }

        /// <summary>
        /// The run properties to read, those of the run this paragraph wraps, its hyperlink's run or its content
        /// control's first run, or <see cref="OfficeXmlNode.None"/> when there are none.
        /// </summary>
        private OfficeXmlNode RunPropertiesNode {
            get {
                if (OwnedModel != null) {
                    return _ownedRunNode == NoNode ? default : new OfficeXmlNode(_ownedTree!, _ownedRunNode).Child(W, "rPr");
                }
                return OfficeXmlNode.FromSdk(SdkScopedRunPropertiesElement);
            }
        }

        /// <summary>The run properties of the hyperlink's run, or of the paragraph's own run, ignoring a content control.</summary>
        private OfficeXmlNode RunPropertiesIgnoringContentControl {
            get {
                if (OwnedModel != null) return RunPropertiesNode;
                return OfficeXmlNode.FromSdk(SdkRunPropertiesIgnoringContentControlElement);
            }
        }

        /// <summary>
        /// The run properties to edit, created as the run's first child when missing, along with the run itself. A
        /// hyperlink's run is used unless the paragraph wraps a content control, or always when
        /// <paramref name="hyperlinkInContentControl"/> is set.
        /// </summary>
        private OfficeXmlNode EditRunProperties(bool hyperlinkInContentControl = false) {
            WordPackageModel? model = OwnedModel;
            if (model == null) return OfficeXmlNode.FromSdk(SdkEditRunPropertiesElement(hyperlinkInContentControl));
            OfficeXmlTree tree = _ownedTree!;
            WordOwnedEdits.PrepareBodyEdit(model);
            int run = EnsureOwnedRun(tree);
            int properties = tree.FindChild(run, W, "rPr");
            if (properties == NoNode) {
                properties = tree.CreateElement(W, "rPr");
                tree.InsertBefore(run, properties, tree.FirstChildOf(run));
            }
            model.MarkChanged();
            return new OfficeXmlNode(tree, properties);
        }

        /// <summary>
        /// Copies a run property the paragraph just set to the runs of a page number field built from it, which only
        /// the SDK path creates.
        /// </summary>
        private void SyncPageNumberFormatting(OfficeXmlNode properties, string localName) {
            if (properties.IsSdk) SdkSyncPageNumberFormatting(properties.SdkElement!, localName);
        }
    }
}
