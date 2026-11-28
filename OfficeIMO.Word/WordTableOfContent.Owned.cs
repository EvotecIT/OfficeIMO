using OfficeIMO.Core.Internal.OpenXml;
using System.Collections.Concurrent;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordTableOfContent {
    private static readonly ConcurrentDictionary<WordTableOfContentsStyle, Lazy<OfficeXmlNode>> Templates = new();
    private OfficeXmlNode _ownedSdt;
    private bool _removed;

    private WordTableOfContent(WordDocument document, OfficeXmlNode control) {
        _document = document;
        _ownedSdt = control;
        Style = WordTableOfContentsStyle.Template1;
        document._ownedModel!.ParagraphsHandedOut = true;
        GetPlaceholderPart();
    }

    private OfficeXmlNode ResolveControl() {
        if (_ownedSdt.IsNone) return default;
        OfficeXmlTree original = _ownedSdt.Tree!;
        if (_document._ownedModel != null) return _ownedSdt;
        object?[]? elements = _document.FindSdkWordNodeMap(original);
        if (elements == null) {
            if (_document.DisposedReadModel is not { } copy) return default;
            if (!ReferenceEquals(original, copy.Document)) {
                int node = _document.DisposedCopyNode(_ownedSdt.Node);
                _ownedSdt = new OfficeXmlNode(copy.Document, node);
            }
            elements = _document.FindSdkWordNodeMap(copy.Document);
            if (elements == null) return _ownedSdt;
        }
        int index = _ownedSdt.Node;
        return index >= 0 && index < elements.Length ? OfficeXmlNode.FromSdk(elements[index]) : default;
    }

    private void RecordOwnedTocEdit() {
        if (_ownedSdt.Exists && _document._ownedModel is { } model) model.MarkChanged();
    }

    private static bool IsTocGallery(OfficeXmlNode block) =>
        block.Child(TocNamespace, "sdtPr").Child(TocNamespace, "docPartObj")
            .Child(TocNamespace, "docPartGallery").Attribute(TocNamespace, "val") == "Table of Contents";

    private static bool IsBlockTocControl(OfficeXmlNode control) {
        const int W = TocNamespace;
        for (OfficeXmlNode parent = control.Parent; parent.Exists; parent = parent.Parent) {
            if (parent.NamespaceId != W) continue;
            if (parent.LocalName is "p" or "tbl" or "tr") return false;
            if (parent.LocalName is "body" or "tc" or "txbxContent") return true;
        }
        return false;
    }

    private OfficeXmlNode GetOwnedGlossaryRoot(WordPackageModel model, bool create) {
        const string relationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "glossaryDocument";
        string? name = WordOwnedSave.RelatedParts(model, relationshipType).FirstOrDefault();
        OfficeXmlTree tree;
        if (name != null) {
            tree = model.Package.GetXml(name);
        } else {
            if (!create || _document.IsDisposed) return default;
            const string contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
            string partName = "/word/glossary/document.xml";
            for (int index = 2; model.Package.PartExists(partName); index++)
                partName = "/word/glossary/document" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".xml";
            _ = model.Package.GetRelationships(model.MainPartName);
            tree = model.Package.CreateXmlPart(partName, contentType, TocNamespace, "glossaryDocument");
            model.Package.AddRelationship(model.MainPartName, relationshipType, partName.Substring("/word/".Length),
                id: "R" + Guid.NewGuid().ToString("N").Substring(0, 16));
            model.MarkChanged();
        }
        return new OfficeXmlNode(tree, tree.Root);
    }

    private static OfficeXmlNode GetOwnedSettingsRoot(WordPackageModel model) {
        const string relationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "settings";
        string? name = WordOwnedSave.RelatedParts(model, relationship).FirstOrDefault();
        if (name == null) return default;
        OfficeXmlTree tree = model.Package.GetXml(name);
        return new OfficeXmlNode(tree, tree.Root);
    }

    internal static bool ReadOwnedUpdateSetting(WordPackageModel model) =>
        WordFieldInventory.ReadFlag(GetOwnedSettingsRoot(model).Child(TocNamespace, "updateFields"), "val");

    internal static bool TrySetOwnedUpdateSetting(WordPackageModel model, bool value) {
        const int W = TocNamespace;
        OfficeXmlNode settings = GetOwnedSettingsRoot(model);
        if (settings.IsNone) return false;
        OfficeXmlNode update = settings.Child(W, "updateFields");
        if (update.IsNone) {
            update = settings.CreateElement(W, "updateFields");
            OfficeXmlNode next = settings.FirstChild;
            while (next.Exists && !IsAfterUpdateSetting(next)) next = next.NextSibling;
            settings.InsertBefore(update, next);
        }
        update.SetAttribute(W, "val", value ? "true" : "false");
        model.MarkChanged();
        return true;
    }

    private static bool IsAfterUpdateSetting(OfficeXmlNode node) =>
        node.NamespaceId == TocNamespace && node.LocalName is
            "hdrShapeDefaults" or "footnotePr" or "endnotePr" or "compat" or "docVars" or "rsids" or "uiCompat97To2003" or
            "attachedSchema" or "themeFontLang" or "clrSchemeMapping" or "doNotIncludeSubdocsInStats" or "doNotAutoCompressPictures" or
            "forceUpgrade" or "captions" or "readModeInkLockDown" or "shapeDefaults" or "decimalSymbol" or "listSeparator" ||
        node.Is(OfficeXmlNamespaces.Math, "mathPr") ||
        node.Is("http://schemas.openxmlformats.org/schemaLibrary/2006/main", "schemaLibrary") ||
        node.NamespaceUri is "http://schemas.microsoft.com/office/word/2010/wordml" or "http://schemas.microsoft.com/office/word/2012/wordml";

    private bool ReadTocUpdateOnOpen() => _document._ownedModel is { } model
        ? ReadOwnedUpdateSetting(model)
        : _document.Settings.UpdateFieldsOnOpen;

    private void EnableTocUpdateOnOpen() {
        if (_document._ownedModel is { } model && TrySetOwnedUpdateSetting(model, true)) return;
        _document.Settings.UpdateFieldsOnOpen = true;
    }

    private bool TryAppendOwnedTemplate(WordTableOfContentsStyle style) {
        if (_document._ownedModel is not { } model || !model.HasSettings ||
            _document.TryUseBuiltInStyleWithoutSdk(model, "TOCHeading") == null) return false;
        const int W = TocNamespace;
        OfficeXmlNode styles = GetTocStylesRoot();
        if (styles.IsNone) return false;
        OfficeXmlNode template = Templates.GetOrAdd(style, value => new Lazy<OfficeXmlNode>(() => {
            byte[] markup = WordEmbeddedMarkup.GetBuildingBlockMarkup(WordEmbeddedMarkup.TablesOfContents, value.ToString(), nameof(style));
            OfficeXmlTree tree = OfficeXmlTree.Parse(markup, markup.Length, new OfficeXmlNamespaces.Table());
            return new OfficeXmlNode(tree, tree.Root);
        })).Value;

        WordOwnedEdits.PrepareBodyEdit(model);
        OfficeXmlNode body = new OfficeXmlNode(model.Document, model.Body);
        _ownedSdt = new OfficeXmlNode(model.Document, model.Document.ImportElement(template.Tree!, template.Node));
        _document.AssignNewSdtIds(_ownedSdt);
        LocalizeOwnedTocStyles(_ownedSdt, model);
        body.InsertBefore(_ownedSdt, body.Child(W, "sectPr"));
        if (WordTableOfContentStyles.EnsureDefined(styles, body)) model.InvalidateOwnedStyles();
        model.ParagraphsHandedOut = true;
        model.NoteBodyFieldAdded();
        model.MarkChanged();
        return true;
    }

    private static void LocalizeOwnedTocStyles(OfficeXmlNode root, WordPackageModel model) {
        const int W = TocNamespace;
        foreach (OfficeXmlNode reference in WordFieldInventory.EnumerateFieldContent(root)) {
            if (!(reference.Is(W, "pStyle") || reference.Is(W, "rStyle") || reference.Is(W, "tblStyle"))) continue;
            if (reference.Attribute(W, "val") is { } id && WordBuiltInStyleNames.IsBuiltInId(id) &&
                model.Styles.FindBuiltInStyleId(id) is { } localized && localized != id)
                reference.SetAttribute(W, "val", localized);
        }
    }

    internal bool BelongsToBody(OfficeXmlNode body) => body.Exists && SdtNode.AncestorOrSelf(TocNamespace, "body") == body;

    internal bool TryPositionOwnedTable(int? targetIndex, out int index) {
        index = -1;
        if (_document._ownedModel == null || _ownedSdt.IsNone || _removed) return false;
        OfficeXmlNode body = _ownedSdt.Parent;
        if (targetIndex is int requested && requested >= 0) {
            OfficeXmlNode reference = GetBodyChild(body, requested);
            if (reference.Exists && reference.NextSibling.Exists && reference != _ownedSdt) {
                _ownedSdt.Remove();
                reference = GetBodyChild(body, requested);
                body.InsertBefore(_ownedSdt, reference);
                RecordOwnedTocEdit();
            }
        }
        foreach (OfficeXmlNode child in body.Children()) {
            index++;
            if (child == _ownedSdt) break;
        }
        return true;
    }

    private static OfficeXmlNode GetBodyChild(OfficeXmlNode body, int index) {
        foreach (OfficeXmlNode child in body.Children()) {
            if (index-- == 0) return child;
        }
        return default;
    }

    internal WordTableOfContent? CopyToOwnedSnapshot(WordDocument snapshot) {
        OfficeXmlNode current = SdtNode;
        if (!current.Parent.Is(TocNamespace, "body")) return null;
        int index = 0;
        for (OfficeXmlNode sibling = current.PreviousSibling; sibling.Exists; sibling = sibling.PreviousSibling) index++;

        OfficeXmlNode copy = GetBodyChild(WordFieldInventory.BodyNode(snapshot), index);
        if (!copy.Is(TocNamespace, "sdt")) return null;
        WordTableOfContent? table = copy.Tree != null ? new WordTableOfContent(snapshot, copy)
            : WrapSdkControl(snapshot, copy);
        if (table != null) CopyGeneratedRefreshStateTo(table);
        return table;
    }

    internal bool TryRemoveOwnedTable() {
        if (_ownedSdt.IsNone || _document._ownedModel == null) return false;
        if (!_removed) {
            _ownedSdt.Remove();
            _removed = true;
            RecordOwnedTocEdit();
        }
        return true;
    }
}
