namespace OfficeIMO.Markdown;

public partial class MarkdownDoc {
    private int _objectTreeBindingDeferralDepth;
    private bool _objectTreeBindingPending;

    internal IDisposable DeferObjectTreeBinding(bool completeBindingOnDispose = true) {
        _objectTreeBindingDeferralDepth++;
        return new ObjectTreeBindingDeferral(this, completeBindingOnDispose);
    }

    internal void MarkObjectTreeBound() {
        _objectTreeBindingPending = false;
    }

    /// <summary>Adds a parser-owned probe block whose temporary object tree is never exposed.</summary>
    internal void AddWithoutObjectTreeBinding(IMarkdownBlock block) {
        if (block == null) {
            throw new ArgumentNullException(nameof(block));
        }

        if (block is IFrontMatterMarkdownBlock frontMatter) {
            _frontMatter = frontMatter;
        } else {
            _blocks.Add(block);
            _lastBlock = block;
        }

        _parseResult = null;
    }

    internal void EnsureObjectTreeBound() {
        if (_objectTreeBindingPending) {
            MarkdownObjectTreeBinder.BindDocument(this);
        }
    }

    private void CompleteObjectTreeBindingDeferral(bool completeBinding) {
        if (_objectTreeBindingDeferralDepth <= 0) {
            throw new InvalidOperationException("No object-tree binding deferral is active.");
        }

        _objectTreeBindingDeferralDepth--;
        if (_objectTreeBindingDeferralDepth == 0 && _objectTreeBindingPending && completeBinding) {
            MarkdownObjectTreeBinder.BindDocument(this);
        }
    }

    private void RequestObjectTreeBinding() {
        if (_objectTreeBindingDeferralDepth > 0) {
            _objectTreeBindingPending = true;
            return;
        }

        MarkdownObjectTreeBinder.BindDocument(this);
    }

    private sealed class ObjectTreeBindingDeferral : IDisposable {
        private MarkdownDoc? _document;
        private readonly bool _completeBinding;

        internal ObjectTreeBindingDeferral(MarkdownDoc document, bool completeBinding) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _completeBinding = completeBinding;
        }

        public void Dispose() {
            var document = _document;
            if (document == null) {
                return;
            }

            _document = null;
            document.CompleteObjectTreeBindingDeferral(_completeBinding);
        }
    }
}
