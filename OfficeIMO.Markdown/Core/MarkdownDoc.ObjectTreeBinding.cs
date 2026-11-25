namespace OfficeIMO.Markdown;

public partial class MarkdownDocument {
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
        ArgumentNullException.ThrowIfNull(block);

        if (block is IFrontMatterMarkdownBlock frontMatter) {
            _frontMatter = frontMatter;
        } else {
            _blocks.Add(block);
            _lastBlock = block;
        }

        ClearParseOwnedState();
    }

    /// <summary>Empties a parser-owned holder document whose blocks were added unbound.</summary>
    internal void ClearParserScratchBlocks() {
        _blocks.Clear();
        _lastBlock = null;
        _frontMatter = null;
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

    // An append to a document that is bound and not deferring binding only needs the new block placed.
    // Binding stays marked pending until the placement finishes, so a child getter that throws part way
    // leaves the next change to bind the whole document.
    private bool TryBindAppendedBlockNow(IMarkdownBlock block) {
        if (_objectTreeBindingDeferralDepth != 0 || _objectTreeBindingPending) {
            return false;
        }

        _objectTreeBindingPending = true;
        if (!MarkdownObjectTreeBinder.TryBindAppendedBlock(this, block)) {
            return false;
        }

        _objectTreeBindingPending = false;
        return true;
    }

    private void RequestObjectTreeBinding() {
        if (_objectTreeBindingDeferralDepth > 0) {
            _objectTreeBindingPending = true;
            return;
        }

        MarkdownObjectTreeBinder.BindDocument(this);
    }

    private sealed class ObjectTreeBindingDeferral : IDisposable {
        private MarkdownDocument? _document;
        private readonly bool _completeBinding;

        internal ObjectTreeBindingDeferral(MarkdownDocument document, bool completeBinding) {
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
