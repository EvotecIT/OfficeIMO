using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

/// <summary>
/// Resolves linked running stories once per read operation without caching across document edits.
/// A captured snapshot can also be activated by deferred rendering callbacks.
/// </summary>
internal sealed class WordHeaderFooterResolutionSnapshot {
    private static readonly AsyncLocal<WordHeaderFooterResolutionSnapshot?> Active = new();
    private readonly WordDocument _document;
    private readonly Dictionary<WordSection, Stories> _stories = new();

    private WordHeaderFooterResolutionSnapshot(WordDocument document, CancellationToken cancellationToken) {
        _document = document;
        WordHeader? defaultHeader = null, firstHeader = null, evenHeader = null;
        WordFooter? defaultFooter = null, firstFooter = null, evenFooter = null;
        foreach (WordSection section in document.Sections) {
            cancellationToken.ThrowIfCancellationRequested();
            const int W = OfficeXmlNamespaces.WordprocessingML;
            foreach (OfficeXmlNode reference in section.SectionNode.Children()) {
                string? type = reference.Attribute(W, "type");
                if (type is not ("default" or "first" or "even")) continue;
                var kind = type == "first" ? WordHeaderFooterType.First : type == "even" ? WordHeaderFooterType.Even : WordHeaderFooterType.Default;
                if (reference.Is(W, "headerReference")) {
                    WordHeader? header = section.PageSetup.GetHeader(kind);
                    if (kind == WordHeaderFooterType.First) firstHeader = header;
                    else if (kind == WordHeaderFooterType.Even) evenHeader = header;
                    else defaultHeader = header;
                } else if (reference.Is(W, "footerReference")) {
                    WordFooter? footer = section.PageSetup.GetFooter(kind);
                    if (kind == WordHeaderFooterType.First) firstFooter = footer;
                    else if (kind == WordHeaderFooterType.Even) evenFooter = footer;
                    else defaultFooter = footer;
                }
            }
            _stories.Add(section, new Stories(defaultHeader, firstHeader, evenHeader, defaultFooter, firstFooter, evenFooter));
        }
    }

    /// <summary>Builds a forward resolution snapshot and activates it for this operation.</summary>
    internal static IDisposable Begin(WordDocument document, CancellationToken cancellationToken = default) =>
        new WordHeaderFooterResolutionSnapshot(document, cancellationToken).Use();

    /// <summary>Gets the active snapshot only when it belongs to the requested document.</summary>
    internal static WordHeaderFooterResolutionSnapshot? Current(WordDocument document) =>
        Active.Value is { } snapshot && ReferenceEquals(snapshot._document, document) ? snapshot : null;

    /// <summary>Temporarily activates this immutable resolution map, restoring any enclosing operation.</summary>
    internal IDisposable Use() {
        var scope = new Scope(Active.Value);
        Active.Value = this;
        return scope;
    }

    internal bool TryGet(WordSection section, out Stories stories) => _stories.TryGetValue(section, out stories);

    internal readonly struct Stories {
        private readonly WordHeader? _defaultHeader, _firstHeader, _evenHeader;
        private readonly WordFooter? _defaultFooter, _firstFooter, _evenFooter;

        internal Stories(WordHeader? defaultHeader, WordHeader? firstHeader, WordHeader? evenHeader,
            WordFooter? defaultFooter, WordFooter? firstFooter, WordFooter? evenFooter) {
            _defaultHeader = defaultHeader; _firstHeader = firstHeader; _evenHeader = evenHeader;
            _defaultFooter = defaultFooter; _firstFooter = firstFooter; _evenFooter = evenFooter;
        }

        internal WordHeader? Header(WordHeaderFooterType kind) => kind == WordHeaderFooterType.First ? _firstHeader
            : kind == WordHeaderFooterType.Even ? _evenHeader : _defaultHeader;
        internal WordFooter? Footer(WordHeaderFooterType kind) => kind == WordHeaderFooterType.First ? _firstFooter
            : kind == WordHeaderFooterType.Even ? _evenFooter : _defaultFooter;
    }

    private sealed class Scope : IDisposable {
        private readonly WordHeaderFooterResolutionSnapshot? _previous;
        private bool _disposed;
        internal Scope(WordHeaderFooterResolutionSnapshot? previous) => _previous = previous;
        public void Dispose() {
            if (_disposed) return;
            Active.Value = _previous;
            _disposed = true;
        }
    }
}
