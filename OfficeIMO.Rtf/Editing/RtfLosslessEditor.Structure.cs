using OfficeIMO.Rtf.Diagnostics;
using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

/// <content>Provides syntax-indexed structural operations for lossless RTF editing.</content>
public sealed partial class RtfLosslessEditor {
    /// <summary>Number of direct syntax nodes inside the root RTF group.</summary>
    public int RootNodeCount => _syntaxTree.Root.Children.Count;

    /// <summary>Inserts a validated raw RTF fragment at a direct root-node index.</summary>
    public void InsertRootRtf(int index, string rtfFragment) {
        if (rtfFragment == null) throw new ArgumentNullException(nameof(rtfFragment));
        if (index < 0 || index > RootNodeCount) throw new ArgumentOutOfRangeException(nameof(index));
        IReadOnlyList<RtfNode> fragmentNodes = ParseFragment(rtfFragment);
        var children = new List<RtfNode>(_syntaxTree.Root.Children);
        children.InsertRange(index, fragmentNodes);
        _syntaxTree = _syntaxTree.WithRoot(new RtfGroup(_syntaxTree.Root.Position, children));
    }

    /// <summary>Inserts a plain paragraph at a direct root-node index.</summary>
    public void InsertRootParagraph(int index, string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        InsertRootRtf(index, @"\pard " + RtfTextEncoding.EncodeText(text) + @"\par");
    }

    /// <summary>Removes direct root syntax nodes and returns the number removed.</summary>
    public int RemoveRootNodes(int index, int count = 1) {
        if (index < 0 || index > RootNodeCount) throw new ArgumentOutOfRangeException(nameof(index));
        if (count < 0 || index + count > RootNodeCount) throw new ArgumentOutOfRangeException(nameof(count));
        if (count == 0) return 0;
        var children = new List<RtfNode>(_syntaxTree.Root.Children);
        children.RemoveRange(index, count);
        _syntaxTree = _syntaxTree.WithRoot(new RtfGroup(_syntaxTree.Root.Position, children));
        return count;
    }

    /// <summary>Moves direct root syntax nodes to their final root-node index.</summary>
    public void MoveRootNodes(int fromIndex, int count, int toIndex) {
        if (fromIndex < 0 || fromIndex >= RootNodeCount) throw new ArgumentOutOfRangeException(nameof(fromIndex));
        if (count <= 0 || fromIndex + count > RootNodeCount) throw new ArgumentOutOfRangeException(nameof(count));
        if (toIndex < 0 || toIndex > RootNodeCount - count) throw new ArgumentOutOfRangeException(nameof(toIndex));
        if (fromIndex == toIndex) return;

        var children = new List<RtfNode>(_syntaxTree.Root.Children);
        List<RtfNode> moving = children.GetRange(fromIndex, count);
        children.RemoveRange(fromIndex, count);
        children.InsertRange(toIndex, moving);
        _syntaxTree = _syntaxTree.WithRoot(new RtfGroup(_syntaxTree.Root.Position, children));
    }

    /// <summary>Replaces the nth picture destination while preserving all unrelated syntax.</summary>
    public bool ReplaceImage(int imageIndex, RtfImage replacement) {
        if (imageIndex < 0) throw new ArgumentOutOfRangeException(nameof(imageIndex));
        if (replacement == null) throw new ArgumentNullException(nameof(replacement));
        RtfGroup replacementGroup = CreatePictureGroup(replacement);
        int currentIndex = 0;
        bool replaced = false;
        RtfGroup root = RewriteDestinationGroup(_syntaxTree.Root, "pict", group => {
            if (currentIndex++ != imageIndex) return group;
            replaced = true;
            return replacementGroup;
        });
        if (replaced) _syntaxTree = _syntaxTree.WithRoot(root);
        return replaced;
    }

    /// <summary>Replaces the content of every matching destination group, including headers and footers.</summary>
    public int ReplaceDestinationContent(string destination, string rtfFragment) {
        if (string.IsNullOrWhiteSpace(destination)) throw new ArgumentException("Destination cannot be empty.", nameof(destination));
        if (rtfFragment == null) throw new ArgumentNullException(nameof(rtfFragment));
        IReadOnlyList<RtfNode> fragmentNodes = ParseFragment(rtfFragment);
        int replacements = 0;
        RtfGroup root = RewriteDestinationGroup(_syntaxTree.Root, destination, group => {
            int destinationIndex = FindDestinationControlIndex(group, destination);
            if (destinationIndex < 0) return group;
            var children = group.Children.Take(destinationIndex + 1).ToList();
            children.AddRange(fragmentNodes);
            replacements++;
            return new RtfGroup(group.Position, children);
        });
        if (replacements > 0) _syntaxTree = _syntaxTree.WithRoot(root);
        return replacements;
    }

    private static IReadOnlyList<RtfNode> ParseFragment(string rtfFragment) {
        RtfSyntaxTree fragment = RtfSyntaxTree.Parse("{" + rtfFragment + "}");
        RtfDiagnostic? error = fragment.Diagnostics.FirstOrDefault(diagnostic =>
            diagnostic.Severity == RtfDiagnosticSeverity.Error || diagnostic.Code == "RTF010" || diagnostic.Code == "RTF011");
        if (error != null) throw new FormatException("Invalid RTF fragment: " + error.Message);
        return fragment.Root.Children;
    }

    private static RtfGroup CreatePictureGroup(RtfImage image) {
        RtfDocument temporary = RtfDocument.Create();
        RtfImage generated = temporary.AddImage(image.Format, image.Data);
        generated.SourceWidth = image.SourceWidth;
        generated.SourceHeight = image.SourceHeight;
        generated.DesiredWidthTwips = image.DesiredWidthTwips;
        generated.DesiredHeightTwips = image.DesiredHeightTwips;
        RtfSyntaxTree syntax = RtfSyntaxTree.Parse(temporary.ToRtf(new RtfWriteOptions { IncludeGenerator = false }));
        return FindDestinationGroup(syntax.Root, "pict") ?? throw new InvalidOperationException("Picture fragment could not be generated.");
    }

    private static RtfGroup? FindDestinationGroup(RtfGroup group, string destination) {
        foreach (RtfNode node in group.Children) {
            if (node is RtfGroup child) {
                if (string.Equals(child.Destination, destination, StringComparison.Ordinal)) return child;
                RtfGroup? nested = FindDestinationGroup(child, destination);
                if (nested != null) return nested;
            }
        }

        return null;
    }

    private static RtfGroup RewriteDestinationGroup(RtfGroup group, string destination, Func<RtfGroup, RtfGroup> replacement) {
        if (string.Equals(group.Destination, destination, StringComparison.Ordinal)) return replacement(group);
        bool changed = false;
        var children = new List<RtfNode>(group.Children.Count);
        foreach (RtfNode node in group.Children) {
            if (node is RtfGroup child) {
                RtfGroup rewritten = RewriteDestinationGroup(child, destination, replacement);
                children.Add(rewritten);
                changed |= !ReferenceEquals(rewritten, child);
            } else {
                children.Add(node);
            }
        }

        return changed ? new RtfGroup(group.Position, children) : group;
    }

    private static int FindDestinationControlIndex(RtfGroup group, string destination) {
        for (int index = 0; index < group.Children.Count; index++) {
            if (group.Children[index] is RtfControlWord control && string.Equals(control.Name, destination, StringComparison.Ordinal)) return index;
        }

        return -1;
    }
}
