using System.Runtime.CompilerServices;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    private static readonly ConditionalWeakTable<OfficeXmlTree, Dictionary<int, List<OfficeXmlNode>>> OwnedPageNumberRuns = new();
    private static readonly ConditionalWeakTable<object, List<OfficeXmlNode>> SdkPageNumberRuns = new();

    private OfficeXmlNode EditPageNumberSourceRun() {
        if (OwnedModel is not { } model) return OfficeXmlNode.FromSdk(VerifyRun());
        WordOwnedEdits.PrepareBodyEdit(model);
        model.MarkChanged();
        return new OfficeXmlNode(_ownedTree!, EnsureOwnedRun(_ownedTree!));
    }

    private static List<OfficeXmlNode>? GetPageNumberRuns(OfficeXmlNode run, bool create) {
        if (run.Tree is { } tree) {
            if (!OwnedPageNumberRuns.TryGetValue(tree, out var bySource)) {
                if (!create) return null;
                bySource = OwnedPageNumberRuns.GetValue(tree, _ => new Dictionary<int, List<OfficeXmlNode>>());
            }
            if (!bySource.TryGetValue(run.Node, out var runs) && create) {
                runs = new List<OfficeXmlNode>();
                bySource.Add(run.Node, runs);
            }
            return runs;
        }
        if (run.SdkElement is not { } element) return null;
        if (create) return SdkPageNumberRuns.GetValue(element, _ => new List<OfficeXmlNode>());
        return SdkPageNumberRuns.TryGetValue(element, out var group) ? group : null;
    }

    private static void InheritPageNumberFormatting(OfficeXmlNode source, OfficeXmlNode previous) {
        List<OfficeXmlNode> runs = GetPageNumberRuns(source, create: true)!;
        OfficeXmlNode sourceProperties = source.Child(W, "rPr");
        for (OfficeXmlNode child = previous.NextSibling; child.Exists; child = child.NextSibling) {
            if (child.Is(W, "r")) {
                AddRun(child);
            } else {
                foreach (OfficeXmlNode run in child.ContentDescendants(W, "r")) AddRun(run);
            }
        }

        void AddRun(OfficeXmlNode run) {
            OfficeXmlNode target = run.GetOrPrependChild(W, "rPr");
            foreach (OfficeXmlNode property in sourceProperties.Children()) {
                RemoveMatchingRunProperties(target, property.NamespaceUri, property.LocalName);
                target.InsertOrdered(property.Clone(), WRPrView.Order);
            }
            runs.Add(run);
        }
    }

    private static void RemoveMatchingRunProperties(OfficeXmlNode properties, string namespaceUri, string localName) {
        for (OfficeXmlNode item = properties.FirstChild; item.Exists;) {
            OfficeXmlNode next = item.NextSibling;
            if (item.Is(namespaceUri, localName)) item.Remove();
            item = next;
        }
    }

    // A formatting change belongs to the authoring run. Runs in unrelated paragraphs
    // (including one moved out of this paragraph by a later edit) are left alone.
    private void SyncPageNumberFormatting(OfficeXmlNode properties, string localName) {
        OfficeXmlNode source = properties.Parent;
        if (GetPageNumberRuns(source, create: false) is not { } runs) return;
        OfficeXmlNode value = properties.Child(W, localName);
        string wordNamespace = OfficeXmlNamespaces.KnownUri(W)!;
        OfficeXmlNode paragraph = ParagraphNode;
        foreach (OfficeXmlNode run in runs) {
            if (!run.IsInside(paragraph)) continue;
            OfficeXmlNode target = run.GetOrPrependChild(W, "rPr");
            RemoveMatchingRunProperties(target, wordNamespace, localName);
            if (value.Exists) target.InsertOrdered(value.Clone(), WRPrView.Order);
        }
    }

    // The live package retains run identity when an interop caller opens the SDK.
    internal static void BindPageNumberFormatting(OfficeXmlTree tree, object?[] elements) {
        if (!OwnedPageNumberRuns.TryGetValue(tree, out var groups)) return;
        foreach (var group in groups) {
            if (group.Key >= elements.Length || elements[group.Key] is not { } source) continue;
            List<OfficeXmlNode> runs = GetPageNumberRuns(OfficeXmlNode.FromSdk(source), create: true)!;
            foreach (OfficeXmlNode run in group.Value) {
                if (run.Node < elements.Length && elements[run.Node] is { } item) runs.Add(OfficeXmlNode.FromSdk(item));
            }
        }
        OwnedPageNumberRuns.Remove(tree);
    }
}
