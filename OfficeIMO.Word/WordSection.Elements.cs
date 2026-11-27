using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordSection {
    private const int SectionNamespace = OfficeXmlNamespaces.WordprocessingML;
    private OfficeXmlNode _ownedSectionNode;
    private SectionProperties? _sdkSectionProperties;

    internal SectionProperties _sectionProperties {
        get {
            EnsureSdkSection();
            return _sdkSectionProperties ?? throw new InvalidOperationException("The section is no longer part of the document.");
        }
        set => _sdkSectionProperties = value;
    }

    internal OfficeXmlNode SectionNode {
        get {
            if (_ownedSectionNode.Exists) {
                if (_document._ownedModel != null) return _ownedSectionNode;
                if (_document.DisposedReadModel is { } copy) {
                    if (!ReferenceEquals(copy.Document, _ownedSectionNode.Tree)) {
                        int node = _document.DisposedCopyNode(_ownedSectionNode.Node);
                        _ownedSectionNode = new OfficeXmlNode(copy.Document, node);
                    }
                    if (_ownedSectionNode.Exists) return _ownedSectionNode;
                }
            }
            return OfficeXmlNode.FromSdk(_sectionProperties);
        }
    }

    internal OfficeXmlNode EditSectionNode() {
        OfficeXmlNode node = SectionNode;
        if (node.Tree != null && _document._ownedModel is { } model) {
            WordOwnedEdits.PrepareBodyEdit(model);
            model.MarkChanged();
        }
        return node;
    }

    internal OfficeXmlNode EnsureSectionChild(string name) {
        OfficeXmlNode section = EditSectionNode();
        OfficeXmlNode child = FindSectionChild(section, name);
        return child.Exists ? child : AddSectionChild(section, name);
    }

    internal static OfficeXmlNode FindSectionChild(OfficeXmlNode section, string name) {
        OfficeXmlNode direct = section.Child(SectionNamespace, name);
        if (direct.Exists) return direct;
        foreach (OfficeXmlNode child in EffectiveSectionChildren(section)) {
            if (child.Is(SectionNamespace, name)) return child;
        }
        return default;
    }

    internal static OfficeXmlNode AddSectionChild(OfficeXmlNode section, string name) {
        const int W = SectionNamespace;
        const int Mc = OfficeXmlNamespaces.MarkupCompatibility;
        OfficeXmlChildOrder order = WSectPrView.Order;
        int desired = order.RankOf(W, name);

        for (OfficeXmlNode parent = section; parent.Exists;) {
            OfficeXmlNode reference = default;
            OfficeXmlNode nested = default;
            foreach (OfficeXmlNode child in parent.Children()) {
                int rank = order.RankOf(child.NamespaceId, child.LocalName);
                if (child.Is(Mc, "AlternateContent")) {
                    OfficeXmlNode branch = WordAlternateContentResolver.SelectBranch(child);
                    int first = int.MaxValue;
                    foreach (OfficeXmlNode selected in EffectiveSectionChildren(branch)) {
                        int selectedRank = order.RankOf(selected.NamespaceId, selected.LocalName);
                        if (selectedRank < 0) continue;
                        first = Math.Min(first, selectedRank);
                        rank = Math.Max(rank, selectedRank);
                    }
                    if (first <= desired && rank > desired) {
                        nested = branch;
                        break;
                    }
                }
                if (rank > desired) {
                    reference = child;
                    break;
                }
                if (rank >= 0) reference = default;
                else if (reference.IsNone) reference = child;
            }
            if (nested.IsNone) return parent.InsertBefore(section.CreateElement(W, name), reference);
            parent = nested;
        }
        throw new InvalidOperationException("Missing section properties.");
    }

    private static IEnumerable<OfficeXmlNode> EffectiveSectionChildren(OfficeXmlNode parent) {
        const int Mc = OfficeXmlNamespaces.MarkupCompatibility;
        OfficeXmlNode child = parent.FirstChild;
        while (child.Exists) {
            if (child.Is(Mc, "AlternateContent")) {
                OfficeXmlNode first = WordAlternateContentResolver.SelectBranch(child).FirstChild;
                if (first.Exists) {
                    child = first;
                    continue;
                }
            } else {
                yield return child;
            }

            while (child.Exists && child.NextSibling.IsNone && child.Parent != parent) child = child.Parent.Parent;
            child = child.NextSibling;
        }
    }

    private WordSection(WordDocument document, OfficeXmlNode section) {
        _document = document;
        _ownedSectionNode = section;
        _margins = new WordMargins(document, this);
        Borders = new WordBorders(document, this);
        _pageSizes = new WordPageSizes(document, this);
    }

    internal static List<WordSection> LoadOwnedSections(WordDocument document, WordPackageModel model) {
        // The SDK initializes missing final properties when it first loads the sections.
        WordOwnedSave.EnsureBodySectionProperties(model);
        var sections = new List<WordSection>();
        OfficeXmlNode body = new OfficeXmlNode(model.Document, model.Body);
        foreach (OfficeXmlNode paragraph in body.Children(SectionNamespace, "p")) {
            OfficeXmlNode section = paragraph.Child(SectionNamespace, "pPr").Child(SectionNamespace, "sectPr");
            if (section.Exists) sections.Add(new WordSection(document, section));
        }
        sections.Add(new WordSection(document, body.Child(SectionNamespace, "sectPr")));
        model.ParagraphsHandedOut = true;
        return sections;
    }

    internal static WordSection AddOwnedSection(WordDocument document, WordPackageModel model, string breakToken) {
        List<WordSection> sections = document.Sections;
        WordSection preceding = sections[sections.Count - 1];
        OfficeXmlNode final = preceding.EditSectionNode();
        OfficeXmlNode body = new OfficeXmlNode(model.Document, model.Body);
        OfficeXmlNode paragraph = body.CreateElement(SectionNamespace, "p");
        OfficeXmlNode properties = paragraph.Append(paragraph.CreateElement(SectionNamespace, "pPr"));
        OfficeXmlNode boundary = properties.Append(properties.CreateElement(SectionNamespace, "sectPr"));
        boundary.SetAttribute(SectionNamespace, "rsidR", WordHeadersAndFooters.GenerateRsid());
        CopySectionProperties(final, boundary);
        body.InsertBefore(paragraph, final);

        preceding._ownedSectionNode = boundary;
        var section = new WordSection(document, final);
        sections.Add(section);
        section.SetBreakType(breakToken);
        OfficeXmlNode pageNumbers = FindSectionChild(final, "pgNumType");
        if (pageNumbers.Exists) pageNumbers.SetAttribute(SectionNamespace, "start", null);
        return section;
    }

    private WordParagraph? AddOwnedSectionParagraph(bool newRun) {
        if (_document._ownedModel is not { } model || _ownedSectionNode.IsNone) return null;
        OfficeXmlNode body = new OfficeXmlNode(model.Document, model.Body);
        OfficeXmlNode boundary = FindOwnedBodyBoundary(body);
        if (boundary.IsNone) return null;
        if (boundary.Is(SectionNamespace, "sectPr") && body.LastChild != boundary) {
            boundary.Remove();
            body.Append(boundary);
        }
        WordParagraph result = WordOwnedParagraphs.AddParagraph(_document, model, newRun, boundary.Node);
        result.Parent = this;
        return result;
    }

    private bool TryReadOwnedSectionParagraphs(out List<WordParagraph> paragraphs) {
        paragraphs = null!;
        if (_document._ownedModel is not { } model || _ownedSectionNode.IsNone) return false;
        List<WordSection> sections = _document.Sections;
        int ordinal = sections.IndexOf(this);
        if (ordinal < 0) return false;
        OfficeXmlNode body = new OfficeXmlNode(model.Document, model.Body);
        OfficeXmlNode first = body.FirstChild;
        if (ordinal > 0) {
            OfficeXmlNode preceding = sections[ordinal - 1].FindOwnedBodyBoundary(body);
            if (!preceding.Is(SectionNamespace, "p")) return false;
            first = preceding.NextSibling;
        }

        OfficeXmlNode stop = default;
        if (ordinal < sections.Count - 1) {
            OfficeXmlNode boundary = FindOwnedBodyBoundary(body);
            if (!boundary.Is(SectionNamespace, "p")) return false;
            stop = boundary.NextSibling;
        }
        return WordOwnedParagraphs.TryCollectRegion(_document, model.Document, first.Node, stop.Node, out paragraphs);
    }

    private OfficeXmlNode FindOwnedBodyBoundary(OfficeXmlNode body) {
        OfficeXmlNode section = SectionNode;
        if (section.Parent == body) return section;
        OfficeXmlNode paragraph = section.Parent.Parent;
        return paragraph.Parent == body && paragraph.Is(SectionNamespace, "p") ? paragraph : default;
    }

    private void EnsureSdkSection() {
        if (_ownedSectionNode.IsNone) return;
        if (_document.IsDisposed) _document.OpenSdkPackageAfterDispose();
        else _document.OpenSdkPackage(null);
    }

    internal void BindSdkSection(WordSection sdk, object?[] nodes) {
        int node = _ownedSectionNode.Node;
        if (_document.IsDisposed && !ReferenceEquals(_ownedSectionNode.Tree, _document._ownedNodeTree)) {
            node = _document.DisposedCopyNode(node);
        }
        if (node < 0 || node >= nodes.Length || !ReferenceEquals(nodes[node], sdk._sdkSectionProperties)) return;
        _ownedSectionNode = default;
        _sdkSectionProperties = sdk._sdkSectionProperties;
        _sdkDocument = sdk._sdkDocument;
        _paragraph = sdk._paragraph;
        _headers.Default = RebindStory(_headers.Default, sdk._headers.Default);
        _headers.First = RebindStory(_headers.First, sdk._headers.First);
        _headers.Even = RebindStory(_headers.Even, sdk._headers.Even);
        _footers.Default = RebindStory(_footers.Default, sdk._footers.Default);
        _footers.First = RebindStory(_footers.First, sdk._footers.First);
        _footers.Even = RebindStory(_footers.Even, sdk._footers.Even);

        static T? RebindStory<T>(T? owned, T? loaded) where T : WordHeaderFooter =>
            owned != null && loaded != null && owned.BindSdkStory(loaded) ? owned : loaded;
    }

    internal bool HasSdkSection => _sdkSectionProperties != null;
}

public partial class WordDocument {
    private List<WordSection>? _ownedSections;

    private void BindOwnedSections() {
        List<WordSection>? owned = _ownedSections;
        if (owned == null || _ownedNodeElements is not { } nodes) return;
        _ownedSections = null;
        if (owned.Count != _sections.Count) return;

        bool allBound = true;
        for (int index = 0; index < _sections.Count; index++) {
            WordSection section = owned[index];
            section.BindSdkSection(_sections[index], nodes);
            if (section.HasSdkSection) _sections[index] = section;
            else allBound = false;
        }
        // Preserve the collection already returned to the caller when the sections correspond.
        if (allBound) _sections = owned;
    }
}
