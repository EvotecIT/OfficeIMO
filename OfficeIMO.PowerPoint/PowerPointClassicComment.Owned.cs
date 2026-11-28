using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointClassicComment {
        private const int PresentationNamespace = OfficeXmlNamespaces.PresentationML;

        private OfficeXmlNode AttachedOwnedComment() {
            if (_ownedComments == null || _presentation.IsSdkLoaded) return default;
            _presentation.ThrowIfDisposedForCommentApi();
            if (_slide.RemovedFromPackage) {
                throw new InvalidOperationException("The comment's slide is no longer attached to the presentation.");
            }
            return _ownedComments.CommentAt(_ownedIndex);
        }

        private bool TryReadOwnedUInt(string attribute, out uint? value) {
            OfficeXmlNode comment = AttachedOwnedComment();
            string? text = comment.Attribute(attribute);
            value = text is { Length: > 0 } ? XmlConvert.ToUInt32(text) : null;
            return !comment.IsNone && text != string.Empty;
        }

        private string? ReadOwnedText() {
            OfficeXmlNode comment = AttachedOwnedComment();
            if (comment.IsNone) return null;
            OfficeXmlNode text = comment.Child(PresentationNamespace, "text");
            if (text.IsNone) return string.Empty;
            OfficeXmlTree tree = text.Tree!;
            int first = tree.FirstChildOf(text.Node);
            if (first == OfficeXmlTree.NoNode) return string.Empty;
            // The SDK reads split text at its first text segment. Let it handle segments and unusual leaf markup.
            if (tree.KindOf(first) != OfficeXmlTreeNodeKind.Text || tree.NextSiblingOf(first) != OfficeXmlTree.NoNode) return null;
            string value = tree.GetText(first);
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        private bool TryReadOwnedCreated(out DateTime? created) {
            OfficeXmlNode comment = AttachedOwnedComment();
            string? text = comment.Attribute("dt");
            created = text is { Length: > 0 } ? XmlConvert.ToDateTime(text, XmlDateTimeSerializationMode.Utc) : null;
            return !comment.IsNone && text != string.Empty;
        }

        private long? ReadOwnedPosition(string axis) {
            OfficeXmlNode position = AttachedOwnedComment().Child(PresentationNamespace, "pos");
            if (position.IsNone) return null;
            string? text = position.Attribute(axis);
            if (text == string.Empty) return null;
            return text == null ? 0L : XmlConvert.ToInt64(text);
        }
    }

    /// <summary>Binds the comments held from one inventory to their corresponding SDK elements in one pass.</summary>
    internal sealed class PowerPointClassicCommentListBinding : IPackageSdkBinding {
        private readonly PowerPointSlide _slide;
        private readonly string _partName;
        private IReadOnlyList<OfficeXmlNode>? _comments;
        private OfficeXmlNode[]? _sdkComments;
        private bool _removed;

        internal PowerPointClassicCommentListBinding(PowerPointDocument document, PowerPointSlide slide, string partName,
            IReadOnlyList<OfficeXmlNode> comments) {
            _slide = slide;
            _partName = partName;
            _comments = comments;
            document.RegisterPackageBinding(this);
        }

        internal OfficeXmlNode CommentAt(int index) => _comments?[index] ?? default;

        internal OfficeXmlNode SdkCommentAt(int index) => _sdkComments?[index] ?? default;

        void IPackageSdkBinding.InvalidateRemoved(bool sdkOpening) => _removed |= _slide.RemovedFromPackage;

        void IPackageSdkBinding.BindSdk() {
            try {
                if (_removed) return;
                OfficeXmlNode[] comments = PowerPointClassicComment.ReadSdkCommentsForBinding(_slide, _partName);
                if (comments.Length == _comments?.Count) _sdkComments = comments;
            } catch (InvalidOperationException) {
                // A wrapper reports when its original comment did not survive the handoff.
            } finally {
                _comments = null;
            }
        }
    }
}
