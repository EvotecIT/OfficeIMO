using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordHeaderFooter {
    private const int RunningStoryNamespace = OfficeXmlNamespaces.WordprocessingML;
    private protected string? _ownedPartName;

    private OfficeXmlTree? ReadOwnedStory(out WordPackageModel? model) {
        model = _document._ownedModel;
        if (_ownedPartName == null || model == null || !model.Package.PartExists(_ownedPartName)) return null;
        return model.Package.GetXml(_ownedPartName);
    }

    internal OfficeXmlNode StoryRoot {
        get {
            if (ReadOwnedStory(out _) is { } tree) return new OfficeXmlNode(tree, tree.Root);
            return OfficeXmlNode.FromSdk((OpenXmlCompositeElement?)_header ?? _footer);
        }
    }

    internal void TrackOwnedParagraphs(OfficeXmlTree tree) => _document.RegisterRunningStory(_ownedPartName!, tree);

    private WordParagraph? AddOwnedStoryParagraph(bool newRun) {
        OfficeXmlTree? tree = ReadOwnedStory(out WordPackageModel? model);
        if (tree == null) return null;
        const int W = RunningStoryNamespace;
        int paragraph = tree.CreateElement(W, "p");
        tree.Append(paragraph, tree.CreateElement(W, "pPr"));
        int run = OfficeXmlTree.NoNode;
        if (newRun) {
            run = tree.CreateElement(W, "r");
            tree.Append(paragraph, run);
        }
        tree.Append(tree.Root, paragraph);
        model!.MarkChanged();
        _document.RegisterRunningStory(_ownedPartName!, tree);
        WordParagraph result = WordParagraph.CreateOwned(_document, tree, paragraph, run);
        result.Parent = this;
        return result;
    }

    private List<WordParagraph>? ReadOwnedStoryParagraphs() {
        OfficeXmlTree? tree = ReadOwnedStory(out _);
        if (tree == null) return null;
        _document.RegisterRunningStory(_ownedPartName!, tree);
        return WordOwnedParagraphs.TryCollectStory(_document, tree, tree.Root, out List<WordParagraph> paragraphs)
            ? paragraphs : null;
    }

    internal bool BindSdkStory(WordHeaderFooter loaded) {
        OpenXmlPartRootElement? root = (OpenXmlPartRootElement?)loaded._headerElement ?? loaded._footerElement;
        if (root?.OpenXmlPart?.Uri.OriginalString != _ownedPartName) return false;
        _headerElement = loaded._headerElement;
        _footerElement = loaded._footerElement;
        return true;
    }

    private OpenXmlPartRootElement? ResolveSdkStory() {
        if (_ownedPartName == null) return null;
        if (_document.IsDisposed && _document._ownedNodeElements == null) _document.OpenSdkPackageAfterDispose();
        _document.OpenSdkPackage(null);
        return _document.FindSdkRunningStory(_ownedPartName);
    }
}

public partial class WordDocument {
    private sealed class RunningStoryBinding {
        internal RunningStoryBinding(string partName) => PartName = partName;
        internal string PartName { get; }
        internal int[]? OriginalPositions { get; set; }
        internal OfficeXmlTree? CopyTree { get; set; }
        internal int[]? CopyNodes { get; set; }
        internal object?[]? SdkElements { get; set; }
        internal object?[]? CopySdkElements { get; set; }
    }

    private Dictionary<OfficeXmlTree, RunningStoryBinding>? _runningStoryBindings;

    internal void RegisterRunningStory(string partName, OfficeXmlTree tree) {
        _runningStoryBindings ??= new Dictionary<OfficeXmlTree, RunningStoryBinding>();
        if (!_runningStoryBindings.ContainsKey(tree)) _runningStoryBindings.Add(tree, new RunningStoryBinding(partName));
    }

    internal bool OwnsWordTree(WordPackageModel model, OfficeXmlTree tree) =>
        ReferenceEquals(tree, model.Document) || _runningStoryBindings?.ContainsKey(tree) == true;

    internal object?[]? FindSdkWordNodeMap(OfficeXmlTree tree) {
        if (ReferenceEquals(tree, _ownedNodeTree)) return _ownedNodeElements;
        if (_runningStoryBindings == null) return null;
        if (_runningStoryBindings.TryGetValue(tree, out RunningStoryBinding? binding)) return binding.SdkElements;
        foreach (RunningStoryBinding story in _runningStoryBindings.Values) {
            if (ReferenceEquals(tree, story.CopyTree)) return story.CopySdkElements;
        }
        return null;
    }

    internal OpenXmlPartRootElement? FindSdkRunningStory(string partName) {
        MainDocumentPart? main = _wordprocessingDocumentValue?.MainDocumentPart;
        if (main == null) return null;
        foreach (IdPartPair pair in main.Parts) {
            if (pair.OpenXmlPart.Uri.OriginalString != partName) continue;
            if (pair.OpenXmlPart is HeaderPart header) return header.Header;
            if (pair.OpenXmlPart is FooterPart footer) return footer.Footer;
        }
        return null;
    }

    private void BindOwnedRunningStories(WordPackageModel model) {
        if (_runningStoryBindings == null) return;
        foreach (var pair in _runningStoryBindings) {
            RunningStoryBinding binding = pair.Value;
            if (FindSdkRunningStory(binding.PartName) is not { } root) continue;
            if (binding.OriginalPositions == null) {
                binding.SdkElements = WordOpenXmlSdkBridge.MapOwnedTree(pair.Key, root);
                WordParagraph.BindPageNumberFormatting(pair.Key, binding.SdkElements);
            } else if (GetDisposedRunningStoryTree(model, pair.Key) is { } copy) {
                object?[] elements = WordOpenXmlSdkBridge.MapOwnedTree(copy, root);
                binding.CopySdkElements = elements;
                int[] positions = binding.OriginalPositions;
                int[] copyNodes = binding.CopyNodes!;
                var mapped = new object?[positions.Length];
                for (int index = 0; index < positions.Length; index++) {
                    int position = positions[index];
                    if (position >= 0 && position < copyNodes.Length) mapped[index] = elements[copyNodes[position]];
                }
                binding.SdkElements = mapped;
                WordParagraph.BindPageNumberFormatting(pair.Key, mapped);
                WordParagraph.BindPageNumberFormatting(copy, elements);
            }
        }
    }

    private void CaptureRunningStoryPositions() {
        if (_runningStoryBindings == null) return;
        foreach (var pair in _runningStoryBindings) {
            pair.Value.OriginalPositions = WordPackageModel.ElementPositions(pair.Key);
        }
    }

    internal OfficeXmlTree? GetDisposedWordTree(WordPackageModel model, OfficeXmlTree source) {
        if (_runningStoryBindings != null) {
            if (_runningStoryBindings.ContainsKey(source)) return GetDisposedRunningStoryTree(model, source);
            foreach (RunningStoryBinding story in _runningStoryBindings.Values) {
                if (ReferenceEquals(story.CopyTree, source)) return source;
            }
        }
        return model.Document;
    }

    private OfficeXmlTree? GetDisposedRunningStoryTree(WordPackageModel model, OfficeXmlTree original) {
        RunningStoryBinding binding = _runningStoryBindings![original];
        if (binding.CopyTree != null) return binding.CopyTree;
        if (!model.Package.PartExists(binding.PartName)) return null;
        OfficeXmlTree tree = model.Package.GetXml(binding.PartName);
        binding.CopyTree = tree;
        binding.CopyNodes = WordPackageModel.ElementsByPosition(tree);
        return tree;
    }

    internal int DisposedCopyNode(OfficeXmlTree tree, int node) {
        if (_runningStoryBindings == null || !_runningStoryBindings.TryGetValue(tree, out RunningStoryBinding? binding))
            return DisposedCopyNode(node);
        if (binding.OriginalPositions is not { } positions || binding.CopyNodes is not { } nodes || node < 0 || node >= positions.Length)
            return OfficeXmlTree.NoNode;
        int position = positions[node];
        return position >= 0 && position < nodes.Length ? nodes[position] : OfficeXmlTree.NoNode;
    }
}
